import json
from pynput.keyboard import Key, Listener
import ctypes
import psutil

# TODO: more precise / convenient save directory
### Settings
# Filenames for main key log and backup file
JSON_FILE_NAME = "key-freq.json"
BACKUP_FILE_NAME = "key-freq-backup.json"

# Save every x number of keystrokes
SAVE_INTERVAL = 500

# print each keypress onto terminal
LOG_TO_TERMINAL = True

LOCKED_IN_GARBAGE_COLLECTION_LIMIT = 5

### Global Variables
KEY_LOG = {}


MODIFIER_KEYS = [
    Key.alt,
    Key.alt_r,
    Key.alt_l,
    Key.cmd,
    Key.cmd_r,
    Key.cmd_l,
    Key.ctrl,
    Key.ctrl_r,
    Key.ctrl_l,
    Key.shift,
    Key.shift_r,
    Key.shift_l,
]

IGNORED_KEYS = []

REMAP = {
    Key.alt_r: Key.alt,
    Key.alt_l: Key.alt,
    Key.ctrl_r: Key.ctrl,
    Key.ctrl_l: Key.ctrl,
    Key.cmd_r: Key.cmd,
    Key.cmd_l: Key.cmd,
    Key.shift_r: Key.shift,
    Key.shift_l: Key.shift,
}

CONTROL_KEYS = {
    "'\\x01'": "a",
    "'\\x02'": "b",
    "'\\x03'": "c",
    "'\\x04'": "d",
    "'\\x05'": "e",
    "'\\x06'": "f",
    "'\\x07'": "g",
    "'\\x08'": "h",
    "'\\x09'": "i",
    "'\\x0a'": "j",
    "'\\x0b'": "k",
    "'\\x0c'": "l",
    "'\\x0d'": "m",
    "'\\x0e'": "n",
    "'\\x0f'": "o",
    "'\\x10'": "p",
    "'\\x11'": "q",
    "'\\x12'": "r",
    "'\\x13'": "s",
    "'\\x14'": "t",
    "'\\x15'": "u",
    "'\\x16'": "v",
    "'\\x17'": "w",
    "'\\x18'": "x",
    "'\\x19'": "y",
    "'\\x1a'": "z",
    "'\\x1b'": "[",
    "'\\x1c'": "\\",
    "'\\x1d'": "]",
    "'\\x1e'": "^",
    "'\\x1f'": "_",
}

keys_currently_down = []
total_keypresses = 0


def load_and_backup_file():
    global KEY_LOG

    try:
        with open(JSON_FILE_NAME, "r") as f:
            KEY_LOG = json.load(f)
            print(f"{JSON_FILE_NAME} loaded")

        with open(BACKUP_FILE_NAME, "w") as b:
            json.dump(KEY_LOG, b, indent=2)
            print(f"Back up made to {BACKUP_FILE_NAME}")

    except Exception as e:
        print("Error loading or backing up data: ", e)
        print("Data will still be saved to a file")


def set_total():
    global KEY_LOG

    attribute = "_TOTAL_LOG"
    KEY_LOG.setdefault(attribute, {})

    for key in KEY_LOG[attribute]:
        KEY_LOG[attribute][key] = 0

    for process_name, log in KEY_LOG.items():
        if not process_name == attribute:
            for key, count in log.items():
                KEY_LOG[attribute].setdefault(key, 0)
                KEY_LOG[attribute][key] += count


def get_process_name():
    user32 = ctypes.WinDLL("user32")
    hwnd = user32.GetForegroundWindow()
    process_id = ctypes.c_ulong()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))

    try:
        process = psutil.Process(process_id.value)
        return process.name()

    except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
        print(f"Error getting process name: {e}")


def record_key(process_name, key):
    global KEY_LOG, total_keypresses

    KEY_LOG.setdefault(process_name, {})
    KEY_LOG[process_name].setdefault(key, 0)
    KEY_LOG[process_name][key] += 1
    total_keypresses += 1

    print(f"{total_keypresses}. {process_name}:  {key} - {KEY_LOG[process_name][key]}")

    # Save Json File every SAVE_INTERVAL keypresses
    if total_keypresses % SAVE_INTERVAL == 0:
        save_data_to_json()


def save_data_to_json():
    global KEY_LOG

    with open(JSON_FILE_NAME, "w") as f:
        json.dump(KEY_LOG, f, indent=2)

    print(f"Key Press Data saved to: {JSON_FILE_NAME}")


### Key Functions
def is_key_a_symbol(key):
    return str(key)[0:4] != "Key."


def key_to_str(key):
    s = str(key)

    if not is_key_a_symbol(key):
        s = f"<{key._name_}>"
    elif s in CONTROL_KEYS:
        s = CONTROL_KEYS[s]
    else:
        s = key.char
        # s = s.encode("latin-1", "backslashreplace").decode("unicode-escape")
        # s = s[1:-1]  # trim the leading and trailing quotes
    return s


def key_down(key):
    if key in keys_currently_down:
        return

    keys_currently_down.append(key)
    record_key(get_process_name(), key_to_str(key))

    if key == Key.f17:
        print("f17, exiting")
        set_total()
        save_data_to_json()
        return False


def key_up(key):
    global keys_currently_down

    try:
        keys_currently_down.remove(key)
    except ValueError:
        if len(keys_currently_down) >= LOCKED_IN_GARBAGE_COLLECTION_LIMIT:
            number_of_modifiers_down = len(
                [k for k in keys_currently_down if k in MODIFIER_KEYS]
            )
            if number_of_modifiers_down == 0:
                keys_currently_down = []


def preprocess_key(key, f):
    k = key

    if key in REMAP:
        k = REMAP[key]

    if k in IGNORED_KEYS:
        return

    return f(k)


def main():
    load_and_backup_file()
    print("Key Freqwincy will begin recording keypresses now")
    with Listener(
        on_press=(lambda key: preprocess_key(key, key_down)),
        on_release=(lambda key: preprocess_key(key, key_up)),
    ) as listener:
        listener.join()


if __name__ == "__main__":
    main()
