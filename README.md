モジュールの使い方

1. Localization.xlsをDropboxからダウンロードし、Localization/フォルダに保存します。
2. モジュールを有効にします。
3. `make drush re:e`を実行します。

スクリプトの使い方：
1. Localization.xlsをDropboxからダウンロードし、Localization/フォルダに保存します。
2. `make sh php`を実行します。
3. `drush scr docroot/modules/custom/risley_export/scripts/export.php`を実行します。

初めて呼び出す際には、かなりのメモリが使用されます。Localization.xlsをダウンロードしない場合、メモリ制限を超える可能性があります。成功するまで実行し続けるだけで十分です。
