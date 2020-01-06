# version2.0

import re
import warnings

import pandas as pd

from .mkdf import DataFrameMaker
from .df_adjuster import latest

# 問題ないので、SettingWithCopyWarningを無視する
warnings.simplefilter('ignore', pd.core.common.SettingWithCopyWarning)


def interlingua_endocytosis(const_path, add_path, save_path):
    """
    引数のpathに対するデータフレーム操作
    const_pathに対してadd_pathを追加し、save_pathに新規保存する
    """
    # 初期値
    sheet_count = None
    file_type = None
    df = []
    data = []
    master_df_list = []

    # 中間言語
    INTER_MN = ('型式', '型番', '商品名', '型式\n', '型番\n', '商品名\n', 'PartName')
    INTER_PN = ('品名', '分類', 'Cmp name')
    INTER_GS = ('図記号', 'Reference')
    INTER_MF = ('メーカー', 'ﾒｰｶｰ', 'Manufacturer')
    INTER_SF = ('主な仕様', '仕様', 'value')
    INTER_PC = ('購入先', '商社', '取引先', '仕入れ先', '仕入先', '仕入れ元', 'サイト名',
                '仕入先名１', 'Vendor')
    INTER_TM = ('納期', '期間', '発送期間', '納入期間',  'Delivery date')
    INTER_UP = ('単価', '部品単価', 'Unit price')
    INTER_QT = ('数量', '数', '購入数', '購入数量', '必要数', '必要数量', 'quantity')
    INTER_ST = ('小計', 'subtotal')
    INTER_DT = ('仕入日',)
    INTER_ALL = {'型式': INTER_MN, '品名': INTER_PN, '図記号': INTER_GS,
                 'メーカー': INTER_MF, '主な仕様': INTER_SF, '購入先': INTER_PC,
                 '納期': INTER_TM, '単価': INTER_UP, '数量': INTER_QT,
                 '仕入日': INTER_DT, '小計': INTER_ST}
    # const_dfに追加する見出し-型番を除いたリスト。この順番で見出しを整列する。
    title_list_rule = ['品名', '図記号', 'メーカー',
                       '主な仕様', '購入先', '納期', '単価', '数量', '小計']
    # title_list_ruleに型番の見出しを追加する位置(.insertの引数)。
    POSITION = 3
    # add_dfからconst_dfに追加しないデータの見出しを選択
    title_remover = ('図記号', '数量', '小計')
    add_rule = title_list_rule[:]

    for remover in title_remover:
        add_rule.remove(remover)

    class CellularDifferentiation(DataFrameMaker):
        """
        取り込んだファイルを必要な形にする
        """

        def __init__(self, file_path):
            """
            正規表現によるファイル種別の判定。
            判定に応じたデータフレーム＋各種インスタンス変数。
            ---
            #インスタンス一覧
            .up_df…df_listの内INTER_MNを含む行よりも上のDataFrame
            .under_df…df_listの内INTER_MNを含む行以下のDataFrame
            .under_dfi…under_dfのindexを検出したINTER_MNに変更したDaraFrame
            .under_dfi_ilist…under_dfiのindexのlist
            .under_dfi_clist…under_dfiのcolumnのlist
            .drop_under_dfi…under_dfiのINTER_MN列に欠損がある行を削除したもの
            .drop_under_dfi_ilsit…drop_under_dfiのindexのlist
            """

            super().__init__(file_path)

            self.up_df = []
            self.under_df = []
            self.under_dfi = []
            self.under_dfi_ilist = []
            self.under_dfi_clist = []
            self.drop_under_dfi = []
            self.drop_under_dfi_ilsit = []
            self.title_list_inter = []

################################################################################
    # データフレームの操作
            for i, df in enumerate(self.df_list):
                title_list_rule_inter = title_list_rule[:]
                match_title_list = title_list_rule[:]
                title_list = []
                for inter_i, inter_str in enumerate(INTER_MN):
                    inter_mn_n = None
                    # pandas.DataFrameに比較演算するとbool型のpandas.DataFrameが得られる
                    # ここでは指定の文字がTureに置き換わったdfが生成される
                    judg_df = self.df_list[i].values == inter_str
                    # INTER_MNのいずれかが何行目にあるのかhit_indexに代入。
                    for index, low in enumerate(judg_df):
                        for value in low:
                            if value == True:
                                hit_index = index
                                inter_mn_n = inter_i
                                break
                        else:
                            continue
                        break
                    else:
                        continue
                    break

                if inter_mn_n is not None:
                    lat_f = True
                    # 処理したdfにINTER_MNが含まれる場合、以下の処理を実行する。
                    # ・INTER_MNが含まれる列（見出し）の配置をlistで保存。
                    # ・INTER_MNが含まれる列（見出し）よりも上のdfをlistで保存。
                    # ・INTER_MNが含まれる列（見出し）から下のdfをlistで保存。
                    # ※各listの配列番号は処理元のdf_listの配列番号と一致する。
                    title_list = list(self.df_list[i].values[hit_index])
                    self.up_df.append(self.df_list[i][:hit_index])
                    self.under_df.append(self.df_list[i][hit_index + 1:])
                    # under_dfのcolumnを見出しにする。
                    self.under_df[i].columns = title_list

                    # 含まれるINTER_MNを組み込んだ、正しい見出しの順番を保存。
                    # 最終的なalignment_under_dfの抽出に使用する。
                    title_list_rule_inter.insert(
                        POSITION, INTER_MN[inter_mn_n])
                    self.title_list_inter.append(title_list_rule_inter)

                    # 既存の見出しを追加する見出しリストから除外
                    for origin_title in title_list:
                        for add_title in match_title_list:
                            if origin_title == add_title:
                                match_title_list.remove(add_title)

                    # 既存を除いた見出しをunder_dfに追加
                    for match_title in match_title_list:
                        self.under_df[i][match_title] = None

                    # under_dfのindexをデフォルトからINTER_MNに変更してlistで保存。
                    self.under_dfi.append(
                        self.under_df[i].set_index(INTER_MN[inter_mn_n]))
                    # under_dfiのインデックスを比較用に素のlistに保存。
                    self.under_dfi_ilist.append(
                        list(self.under_dfi[i].index.values))
                    # interlingua判定で使用するcolumnsのリスト。
                    self.under_dfi_clist.append(
                        list(self.under_dfi[i].columns))
                    # INTER_MNのindexに欠損値がある行を削除したdfのlistを生成。
                    drop_under_df = self.under_df[i].dropna(
                        subset=[INTER_MN[inter_mn_n]])
################################################################################
                    # DataFrameのindexが重複していた場合の処理
                    # under_dfi_clistにINTER_DTを含んでいた場合
                    for column_d in self.under_dfi_clist[i]:
                        for date in INTER_ALL['仕入日']:
                            if column_d == date:
                                drop_under_df = latest(
                                    drop_under_df, date, INTER_MN[inter_mn_n])
                                lat_f = False

                    # under_dfi_clistにINTER_DTを含んでいない場合
                    if lat_f:
                        drop_under_df.drop_duplicates(
                            subset=INTER_MN[inter_mn_n], keep='last',
                            inplace=True)
################################################################################
                    self.drop_under_dfi.append(
                        drop_under_df.set_index(INTER_MN[inter_mn_n]))
                    # under_dfiのインデックスを比較用に素のlistに保存。（欠損値削除）
                    self.drop_under_dfi_ilsit.append(
                        list(self.drop_under_dfi[i].index.values))

                else:
                    # 処理元のdf_listと各listの配列番号を一致させるための処理。
                    self.up_df.append("")
                    self.under_df.append("")
                    self.under_dfi.append("")
                    self.under_dfi_ilist.append("")
                    self.under_dfi_clist.append("")
                    self.drop_under_dfi.append("")
                    self.drop_under_dfi_ilsit.append("")
                    self.title_list_inter.append("")

    def get_key_from_value(d, val):
        """
        listがvalueのdict(d)に対し、各list内の何れかがvalと一致すれば該当keyを返す
        list内にvalが含まれなければNoneを返す
        """
        keys = [k for k, vl in d.items() for v in vl if v == val]
        if keys:
            return keys[0]
        return None
################################################################################

    # 各ファイルのインスタンス生成
    const = CellularDifferentiation(const_path)
    add = CellularDifferentiation(add_path)

    # 型式一致検索とconstデータフレームに対する必要データの統合
    # indexが型番のconst_dfとadd_dfを回す
    for const_i, _ in enumerate(const.under_dfi):
        for add_i, _ in enumerate(add.drop_under_dfi):
            # それぞれのindex(型番)を回す
            for const_ilist_i, const_index in enumerate(const.under_dfi_ilist[const_i]):
                for add_ilist_i, add_index in enumerate(add.drop_under_dfi_ilsit[add_i]):
                    # 比較対象の型番の空白削除
                    const_index_rep = str(const_index)\
                        .replace(" ", "").replace("Ω", "").replace("〃", "")\
                        .replace("　", "").replace("\n", "").replace("(", "")\
                        .replace(")", "")
                    add_index_rep = str(add_index)\
                        .replace(" ", "").replace("Ω", "").replace("〃", "")\
                        .replace("　", "").replace("\n", "").replace("(", "")\
                        .replace(")", "")
                    # 型番のマッチによる判定、大小を区別しない完全一致が含まれていればok
                    if re.search(const_index_rep, add_index_rep, flags=re.IGNORECASE) is not None:
                        # それぞれのcolumn(見出し)を回す
                        for const_column in const.under_dfi_clist[const_i]:
                            for add_column in add.under_dfi_clist[add_i]:
                                # interlingua方式による中間言語の生成
                                add_inter_column = get_key_from_value(
                                    INTER_ALL, add_column)
                                # 中間言語一致確認,add_rule一致,空白value除外,空白見出し除外
                                if const_column == add_inter_column\
                                        and add_inter_column in add_rule\
                                        and add.drop_under_dfi[add_i].loc[add_index, add_column] is not None\
                                        and add_inter_column is not None:
                                    const.under_dfi[const_i].loc[const_index, const_column]\
                                        = add.drop_under_dfi[add_i].loc[add_index, add_column]

################################################################################
    # 完全見出しによるデータフレーム整列抽出と全データフレーム接続と最終出力
    for dfi, _ in enumerate(const.under_dfi):
        if const.under_dfi[dfi] is not "":
            # indexをデフォルトに戻す。
            const.under_dfi[dfi] = const.under_dfi[dfi].reset_index()
            # 正しい順番に見出しを整列したデータフレームの生成。
            alignment_under_df = const.under_dfi[dfi][const.title_list_inter[dfi]]
            # alignment_under_dfのcolumnをintに上書きして、接続時のズレを無くす。
            sh = alignment_under_df.shape
            alignment_under_df.columns = range(sh[1])

            # dfを接続する際に、完全見出しlistからデータフレームを生成し、間にはさむ。
            # 以下を.Tで転置すれば目標の完全見出しデータフレームが生成される。
            heading_df = pd.DataFrame(const.title_list_inter[dfi])

            master_df = pd.concat(
                [const.up_df[dfi], heading_df.T, alignment_under_df])
            master_df_list.append(master_df)

        else:
            master_df_list.append("")

################################################################################
    # データフレーム出力
    # 読み込んだconstデータがシート名を含まない場合、シート名をデフォルトにする
    if const.sheet_name_list is None:
        with pd.ExcelWriter(save_path) as writer:
            for i, sheet in enumerate(master_df_list):
                sheet.to_excel(
                    writer, sheet_name='sheet' + str(i), index=False, header=False)

    # 読み込んだconstデータがシート名を含む場合、シート名をそのまま使う。
    else:
        with pd.ExcelWriter(save_path) as writer:
            for i, sheet in enumerate(master_df_list):
                sheet.to_excel(
                    writer, sheet_name=const.sheet_name_list[i], index=False, header=False)
