from pandas import read_excel


class SolarBattery:
    def __init__(self):
        self.in_path = "SolarBattery/in.xlsx"
        self.out_path = "SolarBattery/out.xlsx"
        self.show()

    def show(self):
        try:
            print("reading data from {}".format(self.in_path))
            df = read_excel(self.in_path)

            print("pasring result")
            for row in [7, 15, 23, 31]:
                for col in range(3, 23):
                    # R
                    df.iloc[row, col] = round(1000 * df.iloc[row-1, col] / df.iloc[row-2, col], 1)
                    # P
                    df.iloc[row+1, col] = round(df.iloc[row-1, col] * df.iloc[row-2, col], 2)
            for idx in range(4):
                p_row = list(df.iloc[8*idx+8, 3:23])
                # Pmax
                df.iloc[40, 3+idx] = max(p_row)
                # Rmax
                df.iloc[35, 3+idx] = df.iloc[8*idx+7, p_row.index(df.iloc[40, 3+idx])+3]
                # Ri
                df.iloc[36, 3+idx] = 1000 * df.iloc[8*idx+3, 3] / int(df.iloc[8*idx+2, 3][:-2])
                # Rmax/Ri
                df.iloc[37, 3+idx] = round(df.iloc[35, 3+idx] / df.iloc[36, 3+idx], 3)
                # U0*IS
                df.iloc[41, 3+idx] = round(df.iloc[8*idx+3, 3] * int(df.iloc[8*idx+2, 3][:-2]), 2)
                # F=Pmax/(U0*IS)
                df.iloc[42, 3+idx] = round(df.iloc[40, 3+idx] / df.iloc[41, 3+idx], 3)
                # Fmean
                df.iloc[44, 3+idx] = round(sum(list(df.iloc[42, 3:7])) / 4, 3)

                df.iloc[35, 3+idx] = round(df.iloc[35, 3+idx], 1)
                df.iloc[36, 3+idx] = round(df.iloc[36, 3+idx], 1)

            print("saving to {}".format(self.out_path))
            df.to_excel(self.out_path, index=False, header=False)
            print("done")
        except Exception as ex:
            print("exception encountered:\n{}".format(ex))
            input()


if __name__ == "__main__":
    show("in.xlsx", "out.xlsx")
