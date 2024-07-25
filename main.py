import subprocess
import time

while True:
    try:
        # Chạy ứng dụng Flask cùng với waitress
        subprocess.run(["python", "-m", "waitress", "--port=82", "app:app"], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Flask gap loi: {e}")
        print("Đang khoi dong flask...")
        time.sleep(1)  # Đợi một khoảng thời gian trước khi khởi động lại
    except Exception as e:
        print(f"Loi khong xac dinh: {e}")
        print("Đang khoi dong lai flask ...")
        time.sleep(1)
