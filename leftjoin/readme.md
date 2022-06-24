# leftjoin.vbs

## ファイル関係フロー図
※ 2ファイルのみの結合に対応しています  
=> source/ にのエクセルファイルを3つ以上格納した場合正常に動作しません  
※ 例として下図ではsource/のファイル名を指定していますが任意です  
=> 別途vbs内のconfig設定を変更してください
```mermaid
flowchart TD
    vbs(leftjoin.vbs) -- "1.実行" --> xlapp[Excelアプリ]
    xlapp -- "2.新規ブックを作成" --> newbook[新規ブック]
    xlapp -- "3.ファイルを開く" --> leftbook(source/leftbook.xlsx)
    leftbook -- "4.シート(1)の内容を複製" --> newbook
    xlapp -- "5.ファイルを開く" --> joinbook(source/joinbook.xlsx)
    joinbook -- "6.シート(1)の内容を複製" --> newbook
    newbook -- "複製シート内の任意項目で紐づけし保存" --> work(workbook.xlsx)
```

## テーブル結合

### SQL
```
SELECT * 
  FROM leftbook
  LEFT JOIN joinbook
  ON leftbook.other-id = joinbook.id;
```

### leftbook
| id | name | address | other-id |
|:-----------|:-----------|:-----------|:------------|
| 1 | Jhon | AAA | 101 |
| 2 | Mike | BBB | 203 |
| 3 | Rin | CCC | 101 |
| 4 | Nick | DDD | 255 |
| 5 | Anya | EEE | 69 |

### joinbook
| id | other-name | remarks |
|:-----------|:-----------|:-----------|
| 69 | XXX | Remark~X |
| 101 | YYY | Remark~Y |
| 203 | ZZZ | Remark~Z |
| 255 | WWW | Remark~W |

### workbook
| id | name | address | other-id | other-name | remarks |
|:-----------|:-----------|:-----------|:------------|:-----------|:-----------|
| 1 | Jhon | AAA | 101 | YYY | Remark~Y |
| 2 | Mike | BBB | 203 | ZZZ | Remark~Z |
| 3 | Rin | CCC | 101 | YYY | Remark~Y |
| 4 | Nick | DDD | 255 | WWW | Remark~W |
| 5 | Anya | EEE | 69 | XXX | Remark~X |
