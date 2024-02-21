使い方

1. Localization.xlsをDropboxからダウンロードし、Localization/フォルダに保存します。
2. モジュールを有効にします。
3. `make sh php` を実行します。
4. `php -d memory_limit=-1 vendor/bin/drush re:e` を実行します。

初めて呼び出す際には、かなりのメモリが使用されます。Localization.xlsをダウンロードしない場合、メモリ制限を超える可能性があります。Localization.xlsを更新した後は、`make drush re:e` を実行するだけで十分です。