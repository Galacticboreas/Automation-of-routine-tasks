import argparse
import pandas as pd

from PySide6.QtCore import QDateTime, QTimeZone


def transform_date(utc, timezone=None):
    utc_fmt = "yyyy-MM-ddTHH:mm:ss.zzzZ"
    new_date = QDateTime().fromString(utc, utc_fmt)
    if timezone:
        new_date.setTimeZone(timezone)
    return new_date


def read_data(fname):
    df = pd.read_csv(fname)

    # Удаляем неправильные величины
    df = df.drop(df[df.mag < 0].index)
    magnitudes = df["mag"]

    # Моя локальная временная зона
    timezone = QTimeZone(b"Europe/Moscow")

    # Преобразуем временную метку в наш часовой пояс
    times = df["time"].apply(lambda x: transform_date(x, timezone))

    return times, magnitudes


if __name__ == "__main__":
    options = argparse.ArgumentParser()
    options.add_argument("-f", "--file", type=str, required=True)
    args = options.parse_args()
    data = read_data(args.file)
    print(data)

# py tutorial/filtering_data_CSV.py -f data/all_hour.csv 

# (0    PySide6.QtCore.QDateTime(2023, 12, 6, 12, 56, ...
# 1    PySide6.QtCore.QDateTime(2023, 12, 6, 12, 53, ...
# 2    PySide6.QtCore.QDateTime(2023, 12, 6, 12, 37, ...
# 3    PySide6.QtCore.QDateTime(2023, 12, 6, 12, 28, ...
# 4    PySide6.QtCore.QDateTime(2023, 12, 6, 12, 16, ...
# 5    PySide6.QtCore.QDateTime(2023, 12, 6, 12, 8, 2...
# 6    PySide6.QtCore.QDateTime(2023, 12, 6, 12, 7, 7...
# Name: time, dtype: object, 0    1.85
# 1    2.50
# 2    2.10
# 3    2.45
# 4    2.64
# 5    2.20
# 6    2.00
# Name: mag, dtype: float64)
