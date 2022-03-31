import pyautogui as ag

ag.PAUSE = 2


def main():
    print("programa iniciado")
    ag.press("winleft")
    ag.write("ejecutar")
    ag.press("enter")
    ag.write(r"W:\Balances\CONSOLIDADO")
    ag.press('enter')
    print("programa terminado")


if __name__ == '__main__':
    main()