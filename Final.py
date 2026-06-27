import win32com.client as win32
import time
import keyboard   # pip install keyboard

# Connect to CANoe
canoe = win32.DispatchEx("CANoe.Application")
canoe.Open(r"C:\Users\SWAPNIL\Desktop\Projects\Indicator_Project\Indicator.cfg")
time.sleep(3)
canoe.Measurement.Start()
time.sleep(3)

def set_SysVar(app, ns_name, sysvar_name, var):
    if app is not None:
        systemCAN = app.System.Namespaces
        sys_namespace = systemCAN(ns_name)
        sys_value = sys_namespace.Variables(sysvar_name)
        sys_value.Value = var
    else:
        raise RuntimeError("CANoe is not open")

def blink_indicator(app, ns_name, sysvar_name, on_time=0.5, off_time=0.5):
    print("Press 'c' to stop blinking...")
    while True:
        if keyboard.is_pressed("c"):
            set_SysVar(app, ns_name, sysvar_name, 0)  # Ensure OFF
            print("Blinking stopped.")
            break

        # Blink ON
        set_SysVar(app, ns_name, sysvar_name, 1)
        time.sleep(on_time)

        # Blink OFF
        set_SysVar(app, ns_name, sysvar_name, 0)
        time.sleep(off_time)

def blink_hazard(app, ns_name, left_var, right_var, on_time=0.5, off_time=0.5):
    print("Hazard blinking started. Press 'c' to stop...")
    while True:
        if keyboard.is_pressed("c"):
            set_SysVar(app, ns_name, left_var, 0)
            set_SysVar(app, ns_name, right_var, 0)
            print("Hazard blinking stopped.")
            break

        # Blink ON (both)
        set_SysVar(app, ns_name, left_var, 1)
        set_SysVar(app, ns_name, right_var, 1)
        time.sleep(on_time)

        # Blink OFF (both)
        set_SysVar(app, ns_name, left_var, 0)
        set_SysVar(app, ns_name, right_var, 0)
        time.sleep(off_time)

print("Press 'l' for LEFT, 'r' for RIGHT, 'z' for HAZARD...")

# Wait until key is pressed
while True:
    if keyboard.is_pressed("l"):
        print("Left indicator blinking started.")
        blink_indicator(canoe, "Indicator", "Sys_BCM_L", on_time=0.5, off_time=0.5)
        break
    if keyboard.is_pressed("r"):
        print("Right indicator blinking started.")
        blink_indicator(canoe, "Indicator", "Sys_BCM_R", on_time=0.5, off_time=0.5)
        break
    if keyboard.is_pressed("z"):
        print("Hazard blinking started.")
        blink_hazard(canoe, "Indicator", "Sys_BCM_L", "Sys_BCM_R", on_time=0.5, off_time=0.5)
        break
time.sleep(3)
canoe.Measurement.Stop()
canoe.Quit()