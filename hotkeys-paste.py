import pandas as pd
import keyboard
import sys
import time

# Read the messages and shortcut keys from the Excel file
df = pd.read_excel('messages.xlsx')

# Convert the dataframe to a dictionary
messages = df.set_index('Shortcut')['Message'].to_dict()

# Define a function to paste a message based on the given key
def paste_message(key):
    if key in messages:
        message = messages[key]
        keyboard.press_and_release('ctrl+v')
        keyboard.write(message)
        print(f'วางข้อความแล้ว : {message}')

# Register the shortcut keys
for key in messages:
    keyboard.add_hotkey(key, paste_message, args=(key,))

# Start the keyboard listener
print('ปุ่มลัดพร้อมทำงาน. กด ctrl+alt+q เพื่อปิด.')
while True:
    if keyboard.is_pressed('ctrl+alt+q'):
        sys.exit()
    else:
        time.sleep(0.01)
