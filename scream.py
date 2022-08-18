import os
import pandas as pd
import numpy as np
from typing import Optional

COMPANY = f"{os.getenv('company')}/"


class ProccessExcel:
    def __init__(self) -> None:
        self.df: pd.DataFrame = pd.read_excel("scream_master_api.xlsx", sheet_name="scream", index_col="Index")
        self.preprocess()
        filt = self.df["turned_off"] == "Yes"
        self.off: pd.DataFrame = self.df[filt][:]
        self.on: pd.DataFrame = self.df[~filt][:]
        self.proces_off()

    def __repr__(self):
        return f"Scream Test Excel, Sheet Name: scream"

    def proces_off(self) -> None:
        # remove str or empty.
        self.off["turn_off_date"] = pd.to_datetime(self.off["turn_off_date"], errors='coerce')
        self.off = self.off.dropna(subset=["turn_off_date"])

        self.off["to_month"] = self.off["turn_off_date"].dt.month
        self.off["to_year"] = self.off["turn_off_date"].dt.year
        self.off = self.off[["to_month", "to_year"]].groupby(["to_month", "to_year"]).value_counts()
        # self.off.to_excel("turned_off.xlsx")

    def preprocess(self) -> None:
        self.df.rename(
            {c: c.lower().replace(" ", "_") for c in self.df.columns},
            axis="columns",
            inplace=True
        )


if __name__ == "__main__":
    xl = ProccessExcel()
    num_left = len(xl.on)
    months_left = num_left / xl.off.mean()

