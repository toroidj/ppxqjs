Paper Plane xUI  QuickJS Script Module

QuickJS ( https://bellard.org/quickjs/ ) を利用して PPx 上で
Javascript を実行するPPx Module です。
WSH の JScript とある程度の互換性を備え、WSH Script Module と
スクリプトを共有しやすくなっています。

WSH 版との性能の比較としては、実行速度が jscript.dll と
jscript9.dll との中間程度であって、最初の初期化に掛かる時間が
多少遅い、となります。
