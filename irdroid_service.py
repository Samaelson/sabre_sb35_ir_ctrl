import win32serviceutil
import win32service
import win32event
import servicemanager
import threading
import time
import json
import queue
import serial
from serial.tools import list_ports
import win32pipe
import win32file
import subprocess

from sabre_sb35_ir_commands import ir_commands

VID = 0x04D8
PID = 0xFD08
PIPE_NAME = r"\\.\pipe\irdroid"


# ---------------- UTIL ----------------
def find_com_port():
    for port in list_ports.comports():
        if port.vid == VID and port.pid == PID:
            return port.device
    return None


# ---------------- SERVICE ----------------
class IRDroidService(win32serviceutil.ServiceFramework):

    _svc_name_ = "IRDroidService"
    _svc_display_name_ = "IRDroid IR Control Service (Enterprise)"
    _svc_description_ = "Enterprise-grade IRDroid Service (Stable IPC + Metrics + Watchdog)"

    def __init__(self, args):
        super().__init__(args)

        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.running = True

        self.cmd_queue = queue.Queue()

        self.events = {}
        self.res_lock = threading.RLock()
        self.lock = threading.RLock()

        self.ser = None
        self.pipe = None

        self.metrics_lock = threading.RLock()
        self.metrics = {
            "ok": 0,
            "fail": 0,
            "unknown": 0,
            "timeout": 0,
            "total": 0
        }

        self.latency_sum = 0.0

        self.last_recovery = 0
        self.recovery_cooldown = 8
        self.recovery_lock = threading.Lock()

    # ---------------- STOP ----------------
    def SvcStop(self):
        self.running = False
        win32event.SetEvent(self.stop_event)

        with self.lock:
            if self.ser:
                try:
                    self.ser.close()
                except:
                    pass

    # ---------------- START ----------------
    def SvcDoRun(self):
        servicemanager.LogInfoMsg("IRDroid Service started")
        self.init_device_with_retry()

        threading.Thread(target=self.worker_loop, daemon=True).start()
        threading.Thread(target=self.cleanup_watchdog, daemon=True).start()
        threading.Thread(target=self.pipe_server, daemon=True).start()

        win32event.WaitForSingleObject(self.stop_event, win32event.INFINITE)

    # ---------------- LOG ----------------
    def log(self, msg, level="INFO"):
        getattr(servicemanager, f"Log{level}Msg")(f"[IRDroid] {msg}")

    # ---------------- DEVICE INIT ----------------
    def init_device(self):
        port = find_com_port()
        if not port:
            raise Exception("Device not found")

        ser = serial.Serial(port, 115200, timeout=2)

        ser.write(b'\x00\x00\x00\x00\x00')
        time.sleep(0.2)

        ser.write(b'v')
        if b"v225" not in ser.read(4):
            raise Exception("Firmware mismatch")

        ser.write(b'n')
        if b"S01" not in ser.read(3):
            raise Exception("Mode fail")

        with self.lock:
            if self.ser:
                try:
                    self.ser.close()
                except:
                    pass
            self.ser = ser

        self.log("Device ready")

    def init_device_with_retry(self):
        for i in range(10):
            try:
                self.init_device()
                return
            except Exception as e:
                self.log(f"Init attempt {i+1} failed: {e}", "Warning")
                time.sleep(2)

        raise Exception("Device init failed permanently")

    # ---------------- RECOVERY ----------------
    def recover_device(self):
        with self.recovery_lock:
            now = time.time()
            if now - self.last_recovery < self.recovery_cooldown:
                return

            self.last_recovery = now
            self.log("Recovery triggered", "Warning")

            try:
                vidpid = "USB\\VID_04D8&PID_FD08"

                subprocess.run([
                    "powershell",
                    "-Command",
                    f"Disable-PnpDevice -InstanceId '{vidpid}' -Confirm:$false"
                ], check=False)

                time.sleep(1.5)

                subprocess.run([
                    "powershell",
                    "-Command",
                    f"Enable-PnpDevice -InstanceId '{vidpid}' -Confirm:$false"
                ], check=False)

                time.sleep(2)

                self.init_device_with_retry()

            except Exception as e:
                self.log(f"Recovery failed: {e}", "Error")

    # ---------------- METRICS ----------------
    def update_metrics(self, status, latency):
        with self.metrics_lock:
            self.metrics["total"] += 1
            if status in self.metrics:
                self.metrics[status] += 1
            self.latency_sum += latency

    # ---------------- IR SEND ----------------
    def send_ir(self, data):
        try:
            with self.lock:
                if not self.ser or not self.ser.is_open:
                    return False

                self.ser.write(b'\x03')
                self.ser.flush()

                if not self.ser.read(1):
                    return False

                for i in range(0, len(data), 62):
                    chunk = data[i:i+62]
                    written = self.ser.write(chunk)

                    if written != len(chunk):
                        self.log("Partial write detected", "Warning")
                        return False

                return True

        except Exception as e:
            self.log(f"send_ir error: {e}", "Error")
            self.recover_device()
            return False

    # ---------------- WORKER ----------------
    def worker_loop(self):
        while self.running:
            try:
                req_id, cmd = self.cmd_queue.get()

                start = time.time()

                if cmd in ir_commands:
                    ok = self.send_ir(ir_commands[cmd])
                    status = "ok" if ok else "fail"
                else:
                    status = "unknown"

                latency = time.time() - start
                self.update_metrics(status, latency)

                with self.res_lock:
                    event = self.events.get(req_id)
                    if event:
                        event["status"] = status
                        event["event"].set()

            except Exception as e:
                self.log(f"worker error: {e}", "Error")

    # ---------------- WATCHDOG ----------------
    def cleanup_watchdog(self):
        while self.running:
            time.sleep(5)
            now = time.time()

            with self.res_lock:
                stale = [
                    k for k, v in self.events.items()
                    if now - v["created"] > 6
                ]

                for k in stale:
                    self.log(f"Cleaning stale request {k}", "Warning")
                    self.events[k]["event"].set()
                    del self.events[k]

    # ---------------- PIPE ----------------
    def pipe_server(self):
        while self.running:
            try:
                pipe = win32pipe.CreateNamedPipe(
                    PIPE_NAME,
                    win32pipe.PIPE_ACCESS_DUPLEX,
                    win32pipe.PIPE_TYPE_MESSAGE |
                    win32pipe.PIPE_READMODE_MESSAGE |
                    win32pipe.PIPE_WAIT,
                    win32pipe.PIPE_UNLIMITED_INSTANCES,
                    65536, 65536, 0, None
                )

                win32pipe.ConnectNamedPipe(pipe, None)

                while self.running:
                    try:
                        _, data = win32file.ReadFile(pipe, 65536)
                    except:
                        break

                    try:
                        msg = json.loads(data.decode("utf-8"))
                    except:
                        continue

                    if msg.get("cmd") == "exit":
                        self.running = False
                        break

                    req_id = msg.get("id")
                    cmd = msg.get("cmd")

                    if req_id is None:
                        continue

                    with self.res_lock:
                        self.events[req_id] = {
                            "event": threading.Event(),
                            "status": None,
                            "created": time.time()
                        }

                    self.cmd_queue.put((req_id, cmd))

                    with self.res_lock:
                        event_obj = self.events.get(req_id)

                    if event_obj:
                        event_obj["event"].wait(timeout=5)

                    with self.res_lock:
                        status = self.events.get(req_id, {}).get("status")
                        self.events.pop(req_id, None)

                    response = json.dumps({
                        "id": req_id,
                        "status": status or "timeout"
                    }).encode("utf-8") + b"\n"

                    win32file.WriteFile(pipe, response)

                win32file.CloseHandle(pipe)

            except Exception as e:
                self.log(f"Pipe error: {e}", "Error")
                try:
                    win32file.CloseHandle(pipe)
                except:
                    pass
                time.sleep(1)


if __name__ == "__main__":
    win32serviceutil.HandleCommandLine(IRDroidService)