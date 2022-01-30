#!/bash/python3

import pyautogui
import pandas
import time

print(r"""\
        ╓░▒▒╦╗
        jÜ¢¢╩╩▒_               r▒▒Ü¢¢Ä¢ÄÄÄ▒╗╓r                     ╓░░▒▒Ä▒⌐
       j¢╩¢▒R╚ºªÜ▒mr╖╓╓.    ºΓ          º╚¢╩¢¢▒╗▒▒╓        ╓    ╖▒▒╝¢¢¢¢¢╩¢¢▒
      ºº╚ªº        j╚╩¢¢Ä=       ╖r▒▒▒▒▒▒m╓º¢╩¢╩¢¢¢¢¢¢▒▒▒Ü¢╩¢¢¢¢¢¢¢╝╩╩¢╩¢╩¢╩░
      ┌¢▒▒Ü░r▒▒▒»╓▒╝¢╩¢ª      ╓▒Ü╩¢¢¢╩¢¢¢¢¢¢▒╝¢╩¢╩¢╩¢╩¢¢¢╝¢╩¢╩¢╩¢╩¢╩¢╩¢╩╩╩¢╩¢░r░═
      ¢╩¢¢╩¢¢¢¢¢¢¢¢╩╩╩▒      ▒Ü¢╩╩¢╩╩╩╩╩¢╩¢╩¢╝¢╩¢░º╝¢╩╩╩¢╩¢╩╩╩¢╩¢╩¢╩╩╩╩╩¢╩¢╩╩¢¢▒R
        «╩¢╩¢╩¢╩╩╩¢╩¢▒      ▒Ö╩╩¢╩¢╩¢╩¢╩¢╩╩╩¢╩¢╩╩Γ  ªºº ¢╩¢╩¢╩¢╩╩╩¢╩¢╩¢╩¢╩╩╩¢╩¢▒
        º▒╝╩¢╩¢╩¢╩╩╩¢▒     ▐Ü╝¢╝¢╩╩╩¢╩¢╩¢╩╩╩¢ª╙º      ╔▒╝╩╩╩¢╩¢╩¢╩╩╩¢╩¢╩¢╩╩╩¢╩¢¢▒
      º¬²ê╩╩¢╩¢╩¢╩¢╩¢░      ¢╩¢╩¢╩¢╩¢╩¢╩¢╩¢╩¢¢▒▒▒¬     ²╩╩¢╩¢╩¢╩¢╩¢╩¢╩¢╩¢╩¢╩¢╩¢╩¢░
        ª╝¢╩¢╩¢╩¢╩¢╩╩¢      j╝¢╩¢╩¢╩¢╩¢╩¢╩╩╩¢╩¢¢¢Γ «Ä▒▒▒╝╝¢╩¢╩¢╩¢╩¢╩¢╩¢╩¢╩╩╩¢╩¢╝¢╓_
         «¢╩¢╩¢╩¢╩¢╩¢╩▒      º╩╩¢╩¢╩¢╩¢╩¢╩╩╩¢╩¢╩¢░▒Ö¢¢¢¢¢╩╩╩¢╩¢╩¢╩¢╩¢╩╩╩¢╩╩╩¢╩¢╩▒╝▒.
          j╝╩╩¢╩¢╩¢╩╩╩¢▒╖      º¢╩╩╩¢╩¢╩¢╩╩╚j▒╝╩¢Ä¢╝¢╝¢╩¢╩¢╩╩╩¢╩¢╩╩¢¢╩▒╚╙ºΓ
            │¢╚╝╩╩¢╩¢╩¢▒╝▒╗         ºººº  ╓▒╝¢╩╩╩╩¢╩¢╩¢▒╚▒╝╩╩╩╩=º
                 ╩¢╩╩╩¢¬    º!r╗╓   ╓╖r   ╙==  j╩  ╙º
                  ╙ºº └       └¢¢¢▒Ä=º        ¢╩ª
""")

excel_data = pandas.read_excel('receiver.xlsx', sheet_name='Receivers')

count = 0

time.sleep(3)

for column in excel_data['Username'].tolist():
    pyautogui.press('esc')
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(1)
    pyautogui.write(str(excel_data['Username'][count]))
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.press('down')
    pyautogui.press('enter')
    pyautogui.write(str(excel_data['Message'][0]))
    pyautogui.press('enter')
    pyautogui.press('esc')
    count = count + 1

print("The script executed successfully.")
