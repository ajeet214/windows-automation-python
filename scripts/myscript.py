# D:\Jobs\myscript.py
from datetime import datetime
from time import sleep


def main():
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{now}] Have a bright day ✨")
    sleep(10)


if __name__ == "__main__":
    main()
