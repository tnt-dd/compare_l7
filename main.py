import os
import pandas as pd
import numpy as np
from typing import Optional

COMPANY = f"{os.getenv('company')}/"


class ProccessExcel:
    def __init__(self, sheet_name: str) -> None:
        self.sheet_name: str = sheet_name
        self.df: pd.DataFrame = pd.read_excel("cov.xlsx", sheet_name=self.sheet_name, index_col="Index")
        self.grouped: Optional[pd.DataFrame] = None
        self.main()

    def __repr__(self):
        return f"Scream Test Excel, Sheet Name: {self.sheet_name}"

    def preprocess(self) -> None:
        self.df.rename(
            {c: c.lower().replace(" ", "_") for c in self.df.columns},
            axis="columns",
            inplace=True
        )
        if "module_or_file" in self.df.columns:
            self.df.rename(
                {"module_or_file": "module"},
                axis="columns",
                inplace=True
            )
        self.df["class"] = self.df["class"].apply(lambda i: np.nan if i == "#!NULL" else i)
        no_packages_filt = self.df["module"].str.startswith(COMPANY)
        self.df = self.df[no_packages_filt]

    def create_main_col(self) -> None:
        # from COMPANY/accounts/notifications => accounts
        self.df["main"] = self.df["module"].str[len(COMPANY):]
        self.df["main"] = self.df["main"].apply(lambda i: i[:i.find("/")])

    def create_grouped(self) -> None:
        self.grouped = self.df[["main", "count"]].groupby(by=["main"]).sum()
        self.grouped.sort_values(by=["count"], ascending=False, inplace=True)

    def main(self) -> None:
        self.preprocess()
        self.create_main_col()
        self.create_grouped()

    @staticmethod
    def is_not_in(df1: pd.DataFrame, df2: pd.DataFrame, is_in=False) -> pd.DataFrame:
        filt = df1.index.isin(df2.index)
        if is_in:
            return df1[filt]

        return df1[~filt]

    def concat(self, df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
        tmp = df1 + self.is_not_in(df2, df1, is_in=True)
        dfs = [tmp, self.is_not_in(df2, df1)]
        return pd.concat(dfs)

    def __add__(self, other) -> pd.DataFrame:
        if self.is_not_in(self.grouped, other.grouped).empty:
            # all keys in self are in other
            df = self.concat(other.grouped, self.grouped)
        elif self.is_not_in(other.grouped, self.grouped).empty:
            # all keys in other are in self
            df = self.concat(other.grouped, self.grouped)
        else:
            # all keys in both
            df = self.grouped + other.grouped

        return df.sort_values(by="count", ascending=False).round()


if __name__ == "__main__":
    xls = {
        "L7": None,
        "Task_L7": None
    }
    for sheet in xls.keys():
        xls[sheet] = ProccessExcel(sheet_name=sheet)

    l7 = xls["L7"]
    t7 = xls["Task_L7"]
    combo = l7 + t7
    # combo.to_excel("combo1.xlsx", sheet_name="combo1")
