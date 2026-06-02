<?php
// JavaScriptから送られてきた生データ（テキスト）を取得
$content = file_get_contents('php://input');
file_put_contents($_GET['name'].".txt", $content);
?>