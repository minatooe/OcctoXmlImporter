## 使い方
- エクセルブックにOcctoXmlImporterをインポートする。(文字化けしたら適当なエディタでSJISにしてみてください)
- 簡単！
## 書式とか
- OcctoXmlImport({xmlフルパス:string},{param:dic})
## paramの項目
設定する値は、全てstringにしてください。
### typeOfPriPlan
対象の計画名。"発販"or"需調"or"連系線"のどれかを指定。
### typeOfSecPlan
上で指定した計画の中の、さらにどの計画を取りたいのか、"発電"or"調達"or"販売"or"需要"or"連系線"のどれかを指定。
### companyCode
- 需調で配下の事業者別に値をとりたい場合……取得対象のエリアコードを含めた事業者コード4ケタ+エリアコード1ケタを指定。
- 事業者、発電所、取引先等ごとではなく、総計の部分の値をとりたい場合……"総計"を指定。
- それ以外の場合……不要。空文字を指定でも可。
### geneBgCode
- 発販計画が取得対象の場合……発電BGコードを指定。
- 需調、連系線の場合……不要。空文字を指定でも可。
### gridCode
- 取得対象が発電計画……系統コードを指定。
- 取得対象が取引or調達計画……相手先のBGコードを指定。
- 取得対象が連系線の場合……申込番号を指定。
- amountOrNotと同時には設定できません。
### amountOrNot
- PriPlanが発販または需調で、指定したBGまたは事業者ごとの合計値を取りたい場合……"合計"を指定。
- それ以外の場合……不要。空文字を指定でも可。
- gridCodeと同時には設定できません。
## 返り値の仕様
要素2個のvariant型配列を返します。エラーの場合は(0)がfalseで(1)にエラーメッセージが、正常に取得できた場合は(0)がtrueで(1)に取得値の入った要素48個のvariant型配列が入ります。
## いちいち連想配列作るのﾒﾝﾄﾞｸｻ
OcctoXmlImporterにはHashMakeFromArrayという、配列から連想配列を作ってくれる関数も入れてあるので、いいなと思ったら使ってみてください。  
HashMakeFromArray(keyArray1,valueArray2)で、0から順に合成して連想配列を返してくれます。  
返り値の仕様はOcctoXmlImportと同じで、(0)にエラーor正常を示す真偽値が、(1)にはエラーメッセージor作ったDicが入ります。
