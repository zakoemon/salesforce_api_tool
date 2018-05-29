try {
    // 「ServerXMLHTTP」オブジェクト生成
    var http = new ActiveXObject("Msxml2.ServerXMLHTTP");
    // 要求初期化
    http.open("POST", "https://login.salesforce.com/services/oauth2/token", false);
    
    http.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
    
    
    var data = "grant_type=password" +
               "&username=" +
               "&password=" +
               "&client_id=" +
               "&client_secret=";
               
    WScript.Echo(data);

    // 要求
    http.send(data);
    // 応答結果表示
    WScript.Echo(http.status + ":" + http.statusText);
    WScript.Echo(http.getAllResponseHeaders());
    WScript.Echo(http.responseText);
} catch (e) {
    // エラーの場合
    WScript.Echo("Error(" + (e.number & 0xFFFF) + "):" + e.message);
}