#!/usr/bin/env python3
"""
Excel SUMIF完全自動化ツール
パス・タブ・列を指定するだけで複雑な集計を瞬時に実行
"""

import pandas as pd
import os
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

class ExcelSUMIFTool:
    def __init__(self):
        self.data = None
        self.result = None
        
    def load_excel_data(self, file_path, sheet_name=None):
        """Excelファイル読み込み"""
        try:
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"ファイルが見つかりません: {file_path}")
            
            # シート名が指定されていない場合は最初のシートを使用
            if sheet_name is None:
                xl_file = pd.ExcelFile(file_path)
                sheet_name = xl_file.sheet_names[0]
                print(f"📋 シート名未指定のため '{sheet_name}' を使用")
            
            # データ読み込み
            self.data = pd.read_excel(file_path, sheet_name=sheet_name)
            
            print(f"✅ データ読み込み完了")
            print(f"   ファイル: {os.path.basename(file_path)}")
            print(f"   シート: {sheet_name}")
            print(f"   データ数: {len(self.data)}行 x {len(self.data.columns)}列")
            print(f"   列名: {list(self.data.columns)}")
            
            return True
            
        except Exception as e:
            print(f"❌ エラー: {e}")
            return False
    
    def simple_sumif(self, condition_column, condition_value, sum_column):
        """
        基本のSUMIF（単一条件）
        
        Parameters:
        - condition_column: 条件列名
        - condition_value: 条件値
        - sum_column: 合計列名
        """
        try:
            # 条件に合致するデータを抽出
            filtered_data = self.data[self.data[condition_column] == condition_value]
            
            # 合計計算
            total = filtered_data[sum_column].sum()
            count = len(filtered_data)
            
            print(f"\n=== SUMIF結果 ===")
            print(f"条件: {condition_column} = '{condition_value}'")
            print(f"対象データ数: {count}行")
            print(f"合計: {total:,}")
            
            return total
            
        except Exception as e:
            print(f"❌ SUMIF実行エラー: {e}")
            return None
    
    def multiple_sumif(self, conditions, sum_column):
        """
        複数条件SUMIF（SUMIFS相当）
        
        Parameters:
        - conditions: 条件辞書 {'列名': '値', '列名2': '値2'}
        - sum_column: 合計列名
        """
        try:
            filtered_data = self.data.copy()
            
            # 複数条件を順次適用
            condition_text = []
            for col, value in conditions.items():
                if isinstance(value, list):
                    filtered_data = filtered_data[filtered_data[col].isin(value)]
                    condition_text.append(f"{col} IN {value}")
                else:
                    filtered_data = filtered_data[filtered_data[col] == value]
                    condition_text.append(f"{col} = '{value}'")
            
            total = filtered_data[sum_column].sum()
            count = len(filtered_data)
            
            print(f"\n=== 複数条件SUMIF結果 ===")
            print(f"条件: {' AND '.join(condition_text)}")
            print(f"対象データ数: {count}行")
            print(f"合計: {total:,}")
            
            return total
            
        except Exception as e:
            print(f"❌ 複数条件SUMIF実行エラー: {e}")
            return None
    
    def group_sumif(self, group_column, sum_column, condition_column=None, condition_value=None):
        """
        グループ別SUMIF集計
        
        Parameters:
        - group_column: グループ化する列
        - sum_column: 合計列
        - condition_column: 条件列（オプション）
        - condition_value: 条件値（オプション）
        """
        try:
            data = self.data.copy()
            
            # 条件フィルタ適用（指定された場合）
            if condition_column and condition_value:
                data = data[data[condition_column] == condition_value]
                condition_text = f" (条件: {condition_column} = '{condition_value}')"
            else:
                condition_text = ""
            
            # グループ別集計
            result = data.groupby(group_column)[sum_column].agg(['sum', 'count', 'mean']).round(2)
            result.columns = ['合計', 'データ数', '平均']
            
            # 構成比計算
            total = result['合計'].sum()
            result['構成比(%)'] = (result['合計'] / total * 100).round(1)
            
            # 並び替え（合計の降順）
            result = result.sort_values('合計', ascending=False)
            
            print(f"\n=== {group_column}別集計{condition_text} ===")
            print(result)
            
            self.result = result
            return result
            
        except Exception as e:
            print(f"❌ グループ別SUMIF実行エラー: {e}")
            return None
    
    def date_range_sumif(self, date_column, sum_column, start_date=None, end_date=None, group_by='month'):
        """
        日付範囲指定SUMIF
        
        Parameters:
        - date_column: 日付列名
        - sum_column: 合計列名
        - start_date: 開始日（'YYYY-MM-DD'形式）
        - end_date: 終了日（'YYYY-MM-DD'形式）
        - group_by: 集計単位（'month', 'quarter', 'year'）
        """
        try:
            data = self.data.copy()
            
            # 日付列の変換
            data[date_column] = pd.to_datetime(data[date_column])
            
            # 日付範囲フィルタ
            if start_date:
                data = data[data[date_column] >= start_date]
            if end_date:
                data = data[data[date_column] <= end_date]
            
            # 期間別グループ化
            if group_by == 'month':
                data['期間'] = data[date_column].dt.to_period('M')
            elif group_by == 'quarter':
                data['期間'] = data[date_column].dt.to_period('Q')
            elif group_by == 'year':
                data['期間'] = data[date_column].dt.to_period('Y')
            elif group_by == 'week':
                data['期間'] = data[date_column].dt.to_period('W')
            
            # 期間別集計
            result = data.groupby('期間')[sum_column].agg(['sum', 'count', 'mean']).round(2)
            result.columns = ['合計', 'データ数', '平均']
            
            # 前期比計算
            result['前期比(%)'] = result['合計'].pct_change() * 100
            result['前期比(%)'] = result['前期比(%)'].round(1)
            
            period_text = f"{start_date or '開始'} ~ {end_date or '終了'}"
            print(f"\n=== {group_by}別時系列集計 ({period_text}) ===")
            print(result)
            
            self.result = result
            return result
            
        except Exception as e:
            print(f"❌ 日付範囲SUMIF実行エラー: {e}")
            return None
    
    def advanced_sumif(self, sum_column, filters=None, group_columns=None):
        """
        高度なSUMIF（複数条件・複数グループ対応）
        
        Parameters:
        - sum_column: 合計列名
        - filters: フィルタ条件辞書
        - group_columns: グループ化列のリスト
        """
        try:
            data = self.data.copy()
            
            # フィルタ適用
            filter_text = []
            if filters:
                for col, condition in filters.items():
                    if isinstance(condition, dict):
                        # 範囲条件 {'>=': 100, '<': 1000}
                        for op, value in condition.items():
                            if op == '>=':
                                data = data[data[col] >= value]
                                filter_text.append(f"{col} >= {value}")
                            elif op == '>':
                                data = data[data[col] > value]
                                filter_text.append(f"{col} > {value}")
                            elif op == '<=':
                                data = data[data[col] <= value]
                                filter_text.append(f"{col} <= {value}")
                            elif op == '<':
                                data = data[data[col] < value]
                                filter_text.append(f"{col} < {value}")
                    elif isinstance(condition, list):
                        # リスト条件
                        data = data[data[col].isin(condition)]
                        filter_text.append(f"{col} IN {condition}")
                    else:
                        # 等価条件
                        data = data[data[col] == condition]
                        filter_text.append(f"{col} = '{condition}'")
            
            # グループ化
            if group_columns:
                if isinstance(group_columns, str):
                    group_columns = [group_columns]
                
                result = data.groupby(group_columns)[sum_column].agg(['sum', 'count', 'mean']).round(2)
                result.columns = ['合計', 'データ数', '平均']
                
                # 構成比計算
                total = result['合計'].sum()
                result['構成比(%)'] = (result['合計'] / total * 100).round(1)
                result = result.sort_values('合計', ascending=False)
            else:
                # 全体集計
                total = data[sum_column].sum()
                count = len(data)
                mean = data[sum_column].mean()
                
                result = pd.DataFrame({
                    '合計': [total],
                    'データ数': [count],
                    '平均': [round(mean, 2)]
                })
            
            condition_display = f" (条件: {' AND '.join(filter_text)})" if filter_text else ""
            group_display = f"{' x '.join(group_columns)}別" if group_columns else "全体"
            
            print(f"\n=== {group_display}高度集計{condition_display} ===")
            print(result)
            
            self.result = result
            return result
            
        except Exception as e:
            print(f"❌ 高度SUMIF実行エラー: {e}")
            return None
    
    def save_result(self, output_path=None, include_original=False):
        """結果をExcelファイルに保存"""
        if self.result is None:
            print("❌ 保存する結果がありません")
            return False
        
        try:
            if output_path is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = f"sumif_result_{timestamp}.xlsx"
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # 集計結果
                self.result.to_excel(writer, sheet_name='集計結果')
                
                # 元データも含める場合
                if include_original and self.data is not None:
                    # データが大きい場合は最初の1000行のみ
                    sample_data = self.data.head(1000)
                    sample_data.to_excel(writer, sheet_name='元データ(サンプル)', index=False)
                
                # サマリー情報
                summary_data = [
                    ['処理日時', datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                    ['元データ行数', len(self.data) if self.data is not None else 0],
                    ['結果行数', len(self.result)],
                    ['出力ファイル', os.path.basename(output_path)]
                ]
                summary_df = pd.DataFrame(summary_data, columns=['項目', '値'])
                summary_df.to_excel(writer, sheet_name='処理サマリー', index=False)
            
            print(f"✅ 結果保存完了: {output_path}")
            return True
            
        except Exception as e:
            print(f"❌ 保存エラー: {e}")
            return False

# ===============================
# 実用的な使用例・テンプレート
# ===============================

def create_sample_data():
    """サンプルデータ作成"""
    print("📊 サンプルデータを作成します...")
    
    # 売上データサンプル
    sales_data = pd.DataFrame({
        '日付': pd.date_range('2024-01-01', periods=100, freq='D'),
        '営業担当': ['田中', '佐藤', '鈴木', '高橋', '山田'] * 20,
        '部署': ['営業1部', '営業2部', '営業1部', '営業2部', '営業3部'] * 20,
        '商品カテゴリ': ['PC', 'ソフトウェア', 'サービス', 'PC', 'ソフトウェア'] * 20,
        '売上金額': np.random.randint(10000, 500000, 100),
        '数量': np.random.randint(1, 10, 100),
        '顧客ランク': ['A', 'B', 'C', 'A', 'B'] * 20
    })
    
    sales_data.to_excel('売上データサンプル.xlsx', index=False)
    print("✅ 売上データサンプル.xlsx を作成しました")

def example_basic_usage():
    """基本的な使用例"""
    print("\n" + "="*50)
    print("📈 基本的なSUMIF使用例")
    print("="*50)
    
    tool = ExcelSUMIFTool()
    
    # データ読み込み
    if tool.load_excel_data('売上データサンプル.xlsx'):
        
        # 1. 単一条件SUMIF
        print("\n【例1】特定営業担当の売上合計")
        tool.simple_sumif('営業担当', '田中', '売上金額')
        
        # 2. 複数条件SUMIF
        print("\n【例2】特定部署・商品カテゴリの売上合計")
        conditions = {'部署': '営業1部', '商品カテゴリ': 'PC'}
        tool.multiple_sumif(conditions, '売上金額')
        
        # 3. グループ別集計
        print("\n【例3】部署別売上集計")
        tool.group_sumif('部署', '売上金額')

def example_advanced_usage():
    """高度な使用例"""
    print("\n" + "="*50)
    print("🚀 高度なSUMIF使用例")
    print("="*50)
    
    tool = ExcelSUMIFTool()
    
    if tool.load_excel_data('売上データサンプル.xlsx'):
        
        # 1. 日付範囲指定集計
        print("\n【例1】月別売上推移")
        tool.date_range_sumif('日付', '売上金額', 
                             start_date='2024-01-01', 
                             end_date='2024-03-31', 
                             group_by='month')
        
        # 2. 複雑な条件での集計
        print("\n【例2】高額取引のランク別分析")
        filters = {
            '売上金額': {'>=': 100000},  # 10万円以上
            '顧客ランク': ['A', 'B']     # Aランク・Bランクのみ
        }
        tool.advanced_sumif('売上金額', filters=filters, group_columns='顧客ランク')
        
        # 結果保存
        tool.save_result('高度集計結果.xlsx', include_original=True)

# ===============================
# 簡単実行用関数
# ===============================

def quick_sumif(file_path, sheet_name, condition_col, condition_val, sum_col):
    """
    最も簡単なSUMIF実行
    
    使用例:
    quick_sumif('売上データ.xlsx', 'Sheet1', '部署', '営業1部', '売上金額')
    """
    tool = ExcelSUMIFTool()
    if tool.load_excel_data(file_path, sheet_name):
        return tool.simple_sumif(condition_col, condition_val, sum_col)
    return None

def quick_group_sumif(file_path, sheet_name, group_col, sum_col, save_result=True):
    """
    グループ別集計の簡単実行
    
    使用例:
    quick_group_sumif('売上データ.xlsx', 'Sheet1', '部署', '売上金額')
    """
    tool = ExcelSUMIFTool()
    if tool.load_excel_data(file_path, sheet_name):
        result = tool.group_sumif(group_col, sum_col)
        if save_result and result is not None:
            tool.save_result()
        return result
    return None

def batch_sumif_analysis(file_path, sheet_name, sum_column):
    """
    一括分析（複数の角度から自動集計）
    """
    print("\n" + "="*50)
    print("🔍 一括SUMIF分析")
    print("="*50)
    
    tool = ExcelSUMIFTool()
    if not tool.load_excel_data(file_path, sheet_name):
        return
    
    # データの列を確認
    columns = tool.data.columns.tolist()
    print(f"\n利用可能な列: {columns}")
    
    # 自動的に様々な角度で分析
    categorical_columns = []
    for col in columns:
        if col != sum_column and tool.data[col].dtype == 'object':
            categorical_columns.append(col)
    
    print(f"\n分析対象列: {categorical_columns}")
    
    # 各カテゴリ列でグループ別集計
    for col in categorical_columns[:3]:  # 最大3つまで
        print(f"\n--- {col}別分析 ---")
        try:
            tool.group_sumif(col, sum_column)
        except:
            print(f"⚠️ {col}の分析をスキップしました")
    
    # 結果保存
    tool.save_result(f'一括分析結果_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')

# メイン実行部分
if __name__ == "__main__":
    print("🚀 Excel SUMIF自動化ツール")
    print("="*50)
    
    # サンプルデータ作成
    create_sample_data()
    
    # 基本使用例
    example_basic_usage()
    
    # 高度使用例
    example_advanced_usage()
    
    print("\n" + "="*50)
    print("✅ 全ての例が完了しました！")
    print("生成されたファイルを確認してください：")
    print("- 売上データサンプル.xlsx")
    print("- sumif_result_[タイムスタンプ].xlsx")
    print("- 高度集計結果.xlsx")
