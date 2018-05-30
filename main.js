var sf_con = sf_con || {};
sf_con = (function(){
    
    var _actObj = {
        fs:{
            obj:new ActiveXObject('Scripting.FileSystemObject')
            ,option:{
                //  オープンモード
                FORREADING:1   // 読み取り専用
                ,FORWRITING:2   // 書き込み専用
                ,FORAPPENDING:8    // 追加書き込み
                //  開くファイルの形式
                ,TRISTATE_TRUE:-1  // Unicode
                ,TRISTATE_FALSE:0  // ASCII
                ,TRISTATE_USEDEFAULT:-2   // システムデフォルト
            }
        }
        ,http:{
            obj:new ActiveXObject('Msxml2.ServerXMLHTTP')
        }
    };

    var _userInfo = (function(){
        
        function _getConf(){
            var _file,_json;
    
            try{
                _file = _actObj.fs.obj.OpenTextFile(
                    'config.json'
                    ,_actObj.fs.option.FORREADING
                    ,true
                    ,_actObj.fs.option.TRISTATE_FALSE
                );
            }catch(e){
                WScript.Echo("Error(" + (e.number & 0xFFFF) + "):" + e.message);
                throw 'error';
            }
    
            _json = _file.ReadAll();
            _file.Close();
            
            return eval('(' + _json + ')');
        }
    
        function getUserString(){
            var conf = _getConf();
    
                var key, str_array = [];
                for(key in conf){
                    str_array.push(key+'='+conf[key]);
                }

            return str_array.join('&');         
        }

        return {
            getUserString:getUserString
        }
    })();

    var _sf_connector = (function(){

        function sf_authenticate(){
            var http = _actObj.http.obj;
    
            http.open(
                'POST'
                ,'https://login.salesforce.com/services/oauth2/token'
                ,false
            );
    
            http.setRequestHeader(
                "Content-Type"
                ,"application/x-www-form-urlencoded"
            );

            try {
                http.send(_userInfo.getUserString());
            }catch(e){
                WScript.Echo("Error(" + (e.number & 0xFFFF) + "):" + e.message);
                throw 'error';
            }finally{

            }
            
            return {
                status:http.status
                ,res_header:http.getAllResponseHeaders()
                ,res_obj:eval('(' + http.responseText + ')')
            }  
        };

        return {
            getConnection:sf_authenticate
        }
    })();

    return {
        getConnection:_sf_connector.getConnection
        ,releaseObj:function(){
            var key;
            for(key in _actObj){
                _actObj[key].obj = null;
            }
            WScript.Echo('Released ActiveXObject');
        }
    }

})();

var result = sf_con.getConnection();
WScript.Echo('status: ' + result.status);
WScript.Echo('header: ' + result.res_header);
WScript.Echo('body: ' + result.res_obj);
sf_con.releaseObj();

