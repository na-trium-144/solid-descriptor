# SolidDescriptor

SOLIDWORKSのアセンブリの内容(含んでいる部品、合致 だけあればいいかな?)をxmlにダンプし、xmlに書かれている内容からアセンブリを再構築することで、git等でcadデータをマージできるようにする

現状
* アセンブリを開いてXMLExportのmain()を実行でxmlをエクスポート
* アセンブリを開いてXMLImportのmain()を実行で同名のxmlからインポート
  * 合致が壊れます
