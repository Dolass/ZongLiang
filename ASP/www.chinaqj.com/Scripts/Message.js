var getArgs=(function(){
    var sc=document.getElementsByTagName('script');
    var paramsArr=sc[sc.length-1].src.split('?')[1].split('&');
    var args={},argsStr=[],param,t,name,value;
    for(var i=0,len=paramsArr.length;i<len;i++){
        param=paramsArr[i].split('=');
        name=param[0],value=param[1];
        if(typeof args[name]=="undefined"){
            args[name]=value;
        }
        else if(typeof args[name]=="string"){
            args[name]=[args[name]]
            args[name].push(value);
        }
        else{
            args[name].push(value);
        }
    }
    var showArg=function(x){
        if(typeof(x)=="string"&&!/\d+/.test(x)) return "'"+x+"'";
        if(x instanceof Array) return "["+x+"]"
        return x;
    }
    args.toString=function(){
        for(var i in args) argsStr.push(i+':'+showArg(args[i]));
        return '{'+argsStr.join(',')+'}';
    }
    return function(){
        return args;
    }
}
)();
var owner = "www.chinaqj.com";var sf_mess_cfg = {theme:getArgs()["SkinA"],color:getArgs()["SkinB"],title:"�ͻ�����������ѯ",send:"������ѯ",copyright:"www.chinaqj.com",mbpos:"RD"};var sf_mess_msg = {emailErr: '����д��ȷ�ĵ��������ַ��',messErr: '�������������ѳ������ƣ��뱣����1000�������ڣ�',prefix: '����д',success: '�����Ѿ��յ��������ԣ��Ժ��������ϵ��лл��',fail: '�����Ѿ��յ��������ԣ��Ժ��������ϵ��лл��'};var sf_mess_cols = [{type:"textarea",mbtype: "message",tip: "��ѯ����������",innertip: "�������Ҫ�������ǵ��κβ�Ʒ������ǵĲ�Ʒ���κ����ʣ������ԡ����ǻἰʱ��ϵ����",idname: "content"},{type:"text",mbtype: "tel",tip: "�ֻ�����",innertip: "�����������ֻ�����",idname: "phone"},{type:"text",mbtype: "email",tip: "��������",innertip: "���������ĵ�������",idname: "email"},{type:"text",mbtype: "address",tip: "��ϵ��ַ",innertip: "������������ϵ��ַ",idname: "addr"}];document.write('<script src="../scripts/entry.js" type="text/javascript"></script>');