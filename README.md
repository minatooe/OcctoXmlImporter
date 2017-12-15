## 使い方
- エクセルブックにOcctoXmlImporterをインポートする。
- 簡単！
## 書式とか
- OcctoXmlImport({xmlフルパス:string},{param:dic})
## paramの項目
### typeOfPriPlan
対象の計画名。指定できるのは"発販"or"需調"or"連系線"のみ。現在は発販のみ対応。そのうち他のもできるようにします。
### typeOfSecPlan
上で指定した計画の中の、さらにどの計画を取りたいのか。指定できるのは"発電"or"調達"or"販売"or"需要"のみ。
### companyCode
取得対象のエリアコードを含めた事業者コード4ケタ+エリアコード1ケタ。例えば、00001みたいな。typeOfPriPlanが発販と連系線の場合は不要なので空。
### bgCode
取得対象のBGコード。typeOfPriPlanが需調の場合は不要なので空。
### gridCode
取得対象の系統コード。typeOfPriPlanが連系線の場合は申込番号、typeOfSecPlanが調達or販売の場合は対象の取引先BGコード
### amountOrNot
合計or空。合計値を取りたいのか個別値を取りたいのかを設定する項目です。
## 返り値の仕様
要素2個のvariant型配列を返します。エラーの場合は(0)がfalseで(1)にエラーメッセージが、正常に取得できた場合は(0)がtrueで(1)に取得値の入った要素48個のvariant型配列が入ります。
## いちいち連想配列作るのﾒﾝﾄﾞｸｻ
OcctoXmlImporterにはHashMakeFromArrayという、配列から連想配列を作ってくれる関数も入れてあるので、いいなと思ったら使ってみてください。  
HashMakeFromArray(keyArray1,valueArray2)で、0から順に合成して連想配列を返してくれます。  
返り値の仕様はOcctoXmlImportと同じで、(0)にエラーor正常を示すbooleanが、(1)にはエラーメッセージor作ったDicが入ります。
