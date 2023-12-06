import argparse
import pandas as pd


def read_data(fname):
    return pd.read_csv(fname)


if __name__ == "__main__":
    options = argparse.ArgumentParser()
    options.add_argument("-f", "--file", type=str, required=True)
    args = options.parse_args()
    data = read_data(args.file)
    print(data)
# py tutorial/reading_data_from_CSV.py -f data/all_hour.csv 

#                        time   latitude   longitude  depth   mag magType   nst    gap  ...        type  horizontalError depthError magError magNst     status locationSource  magSource
# 0  2023-12-06T12:56:23.160Z  37.609833 -122.383331  11.86  1.85      md  30.0   80.0  ...  earthquake             0.19       0.37  0.16000   23.0  automatic             nc         nc
# 1  2023-12-06T12:53:47.421Z  61.669700 -148.523200  31.40  2.50      ml   NaN    NaN  ...  earthquake              NaN       0.10      NaN    NaN  automatic             ak         ak
# 2  2023-12-06T12:37:28.903Z  61.889600 -149.441900  31.40  2.10      ml   NaN    NaN  ...  earthquake              NaN       0.20      NaN    NaN  automatic             ak         ak
# 3  2023-12-06T12:28:15.250Z  37.399666 -121.762833   3.32  2.45      md  60.0   29.0  ...  earthquake             0.12       0.42  0.15000   71.0  automatic             nc         nc
# 4  2023-12-06T12:16:12.890Z  17.902500  -66.851667  14.83  2.64      md  18.0  207.0  ...  earthquake             0.58       0.60  0.10296    7.0   reviewed             pr         pr
# 5  2023-12-06T12:08:25.339Z  59.292400 -152.097700  63.40  2.20      ml   NaN    NaN  ...  earthquake              NaN       1.30      NaN    NaN  automatic             ak         ak
# 6  2023-12-06T12:07:07.995Z  62.860100 -150.634800   9.20  2.00      ml   NaN    NaN  ...  earthquake              NaN       0.20      NaN    NaN  automatic             ak         ak
