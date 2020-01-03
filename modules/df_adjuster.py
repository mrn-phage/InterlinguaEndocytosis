import pandas as pd
import os


def latest(df, date, name):
    """
    df = pandasで生成したデータフレーム(DataFrame)
    date = 日付見出し(string)
    name = 型番見出し(string)
    df中のdateが大きいのものを昇順（新しい日付が一番上）にする。
    nameが重複したら一番上のデータを残して削除する。
    """

    # 指定columnsを昇順でソートしデータフレームを上書き
    df.sort_values(by=date, ascending=True, inplace=True)
    # 指定columnsの内、重複したindexの最終行を残して削除しデータフレームを上書き。
    df.drop_duplicates(subset=name, keep='last', inplace=True)
    return df
