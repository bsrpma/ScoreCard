import os
import threading
import time

loading_done = False

def clear():
    os.system('cls' if os.name == 'nt' else 'clear')

def loading_animation():
    dots = ["", ".", "..", "..."]
    i = 0
    while not loading_done:
        clear()
        print(f"loading{dots[i % len(dots)]}")
        time.sleep(0.5)
        i += 1
    clear()
    print("Selesai!")

def long_process():
    time.sleep(5)  # Simulasi proses 5 detik

if __name__ == "__main__":
    t = threading.Thread(target=loading_animation)
    t.start()

    long_process()  # proses utama

    loading_done = True
    t.join()
