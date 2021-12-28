//1、变量 gr_InstallPath 等号后面的参数是插件安装文件的所在的网站目录，一般从网站的根目
//   录开始寻址，插件安装文件一定要存在于指定目录下。
//2、变量 gr_plugin_setup_url 指定WEB报表插件的安装程序下载URL。如果插件创建不成功，将提示用户从此URL下载
//   开发者可以将 gr_plugin_setup_url 的值改为自己服务器的URL，方便用户从便捷的WEB服务器下载WEB报表插件安装程序
//3、gr_Version 等号后面的参数是插件安装包的版本号，如果有新版本插件安装包，应上传新版
//   本插件安装文件到网站对应目录，并更新这里的版本号。
//4、更多详细信息请参考帮助中“报表插件(WEB报表)->在服务器部署插件安装包”部分
var gr_InstallPath = "grinstall", //实际项目中应该写从根目录寻址的目录，如gr_InstallPath="/myapp/report/grinstall"; 
    gr_plugin_setup_url = "http://www.rubylong.cn/download/grbsctl6.exe", //WEB报表客户端安装程序的下载URL，官方网站URL
    //gr_plugin_setup_url = gr_InstallPath + "/grbsctl6.exe", //WEB报表客户端安装程序的下载URL
    gr_plugin_required_url = "http://www.rubylong.cn/gridreport/doc/plugins_browser.htm",
    gr_Version = "6,8,21,0901";

//以下注册号为本机开发测试注册号，报表访问地址为localhost时可以去掉试用标志
//购买注册后，请用您的注册用户名与注册号替换下面变量中值
var gr_UserName = '锐浪报表插件本机开发测试注册',
    gr_SerialNo = '8PJH495VA61FLI5TG0L4KB2337F1G7AKLD6LNNA9F9T28IKRU6N33P8Z6XX4BUYB5E9NZ6INMD5T8EN47IX63VV7F9BJHB5ZJQQ6MX3J3V12C4XDHU97SXX6X3VA57KCB6';

//报表插件目前只能在32位浏览器中使用
var _gr_platform = window.navigator.platform,
    _gr_isX64 = (_gr_platform.indexOf("64") > 0),
    _gr_agent = navigator.userAgent.toLowerCase(),
    _gr_isIE = (_gr_agent.indexOf("msie") > 0),
    gr_CodeBase = _gr_isIE ? 'codebase="' + gr_InstallPath + (_gr_isX64 ? '/grbsctl6x64.cab' : '/grbsctl6.cab') + '#Version=' + gr_Version + '"' : ""; //区分浏览器(IE or not)

var ajaxMode = 1; //指定获取报表模板与报表数据是否采用 ajax 方式。推荐采用ajax方式

//////////////////////////////////////////////////////////////////////////
function MsgPluginFailed() {
    var body = document.body;
    children = body.children,
    referNode = children.length ? children[0] : null,
    newNode1 = document.createElement("h3");
    newNode2 = document.createElement("h3");

    //弹出alert提示信息，可修改为更适合的表述
    alert("创建插件失败，当前浏览器不支持插件，或插件在当前电脑上没有安装！");

    //在网页最前面加上提示下载的文字，可修改为更适合的表述与界面形式
    newNode1.innerHTML = '特别提示：<a href="' + gr_plugin_setup_url + '">点击下载WEB报表插件安装程序</a>，下载后双击下载的文件进行安装，安装完成后重新打开当前网页。';
    newNode2.innerHTML = '请参考<a href="' + gr_plugin_required_url + '" target="_blank">浏览器插件兼容性说明</a>，选用支持插件的浏览器查看当前网页。';
    document.body.insertBefore(newNode1, referNode);
    document.body.insertBefore(newNode2, referNode);
}

//创建报表对象，报表对象是不可见的对象，详细请查看帮助中的 IGridppReport
//Name - 指定插件对象的ID，可以用js代码 document.getElementById("%Name%") 获取报表对象
//EventParams - 指定报表对象的需要响应的事件，如："<param name='OnInitialize' value=OnInitialize> <param name='OnProcessBegin' value=OnProcessBegin>"形式，可以指定多个事件
function CreateReport(PluginID, EventParams) {
    var typeid,
        report;

    if (_gr_isIE)
        typeid = 'classid="clsid:396841CC-FC0F-4989-8182-EBA06AA8CA2F" ';
    else
        typeid = 'type="application/x-grplugin6-report" ';
    typeid += gr_CodeBase;
    document.write('<object id="' + PluginID + '" ' + typeid);

    //报表引擎对象为不可见的对象，将其display样式设置为none应该是最合适的，
    //但如此设置之后,360极速浏览器中报表的方法就不能用，所以按“display:block;margin-top:-16;”设置样式
    //“margin-top:-16;”是为了让报表看起来不占用空间。不然页面上就会出现多余的空白区域
    //document.write(' width="0" height="0" style="display:none;" VIEWASTEXT>');
    document.write(' width="0" height="0" style="display:block;margin-top:-16;" VIEWASTEXT>');

    if (EventParams != undefined)
        document.write(EventParams);
    document.write('</object>');

    document.write('<script type="text/javascript">');
    document.write(PluginID + '.Register("' + gr_UserName + '", "' + gr_SerialNo + '");');
    document.write('</script>');

    report = document.getElementById(PluginID);
    if (!report || !report.Register) {
        MsgPluginFailed();
        return;
    }

    report.Register(gr_UserName, gr_SerialNo);
}

//参数 args 为一个对象，其下成员数据项为创建插件的相关参数，对应成员未定义采用默认值，以下为各个成员数据的简要说明：
//type：指定插件的类型名称，必须为“printviewer”、“displayviewer”与“designer”这三个值之一。
//id：指定插件在网页元素中的 id 值，查询显示器(DisplayViewer)与打印显示器(PrintViewer)的默认值为“ReportViewer”，报表设计器(Designer)的默认值为“ReportDesigner”。
//width：插件的显示宽度，“100%”为整个显示区域宽度，“500px”表示500个屏幕像素点。默认值为“100%”。
//height：插件的显示宽度，“100%”为整个显示区域高度，“500px”表示500个屏幕像素点。默认值为“100%”。
//report：获取报表模板的URL，或报表模板字符串数据。
//data：获取报表数据的URL，JSON(或XML)文本字符串，或JSON对象。
//dataUrlParams：获取报表数据的URL的附加的HTTP请求参数。
//saveurl：指定进行报表设计结果保存的URL，仅报表设计器插件用到。
//autorun：指定插件在创建之后是否自动生成并展现报表，报表设计器插件不会用到。默认为“true”。
//exparams：指定更多的插件属性参数,形如: "<param name="%ParamName%" value="%Value%">"这样的字符串，具体需要查看帮助中对应插件的API。
//oncreate：指定插件在网页中创建后需要执行的回调(事件)函数，通常可以在此回调函数中设置报表对象的响应事件。
//onreportload：指定报表模板已经加载后的回调(事件)函数，通常可以在此回调函数获取报表模板定义相关的数据，并执行相应的任务。
//ondataload：指定报表数据已经加载后的回调(事件)函数。设计器控件不会触发本事件函数。
function doInsertPlugin(args) {
    var type = args.type || "printviewer",
        id = args.id,
        width = args.width || "100%",
        height = args.height || "100%",
        report = args.report || "",
        data = args.data || "",
        dataUrlParams = args.dataUrlParams || "",
        saveurl = args.saveurl || "",
        autorun = args.autorun || (args.autorun == undefined),
        exparams = args.exparams,

        oncreate = args.oncreate,

        isdesigner = (args.type == "designer"),
        _autorun = autorun && !ajaxMode,
        typeid,
        plugin;

    function _viewerStart(viewer, args) {
        var reportobj = viewer.Report,
            torun = args.autorun || (args.autorun == undefined),
            report = args.report || viewer.ReportURL,
            data = args.data || viewer.DataURL,
            onreportload = args.onreportload,
            ondataload = args.ondataload;

        function run() {
            var _report = viewer._gr_report,
                _data = viewer._gr_data;

            if (_report && _data) {
                if (_report != "{}") {
                    reportobj.LoadFromStr(_report);
                    onreportload && onreportload(reportobj);
                }

                if (!reportobj.DataLoaded) {
                    reportobj.LoadDataFromXML(_data);
                }

                ondataload && ondataload(reportobj);

                torun && viewer.Start();

                viewer._gr_report = undefined;
                viewer._gr_data = undefined;
            }
        }

        if (report) {
            if (grplugin_is_url(report)) {
                grplugin_http(report, function (xmlhttp) {
                    viewer._gr_report = xmlhttp.responseText;
                    run();
                });
            }
            else {
                viewer._gr_report = (typeof report == "object") ? JSON.stringify(report) : report;;
                run();
            }
        }
        else {
            viewer._gr_report = "{}";
            run();
        }

        if (data) {
            if (grplugin_is_url(data)) {
                grplugin_http(data, function (xmlhttp) {
                    viewer._gr_data = xmlhttp.responseText;
                    run();
                }, true, args.dataUrlParams);
            }
            else {
                viewer._gr_data = (typeof data == "object") ? JSON.stringify(data) : data;;
                run();
            }
        }
        else {
            viewer._gr_data = "{}";
            run();
        }
    }

    if (isdesigner) {
        id = id || "ReportDesigner";
    }
    else {
        id = id || "ReportViewer";
    }

    if (_gr_isIE) {
        if (isdesigner) {
            typeid = "CE666189-5D7C-42ee-AAA4-E5CB375ED3C7";
        }
        else {
            if (type == "printviewer") {
                typeid = "ABB64AAC-D7E8-4733-B052-1B141C92F3CE";
            }
            else {
                typeid = "600CD6D9-EBE1-42cb-B8DF-DFB81977122E";
            }
        }
        typeid = 'classid="clsid:' + typeid + gr_CodeBase
    }
    else {
        typeid = 'type="application/x-grplugin6-' + type + '"';
    }

    document.write('<object id="' + id + '" ' + typeid);
    document.write(' width="' + width + '" height="' + height + '">');
    if (isdesigner) {
        if (!ajaxMode) {
            document.write('<param name="LoadReportURL" value="' + report + '">');
            document.write('<param name="DataURL" value="' + data + '">');
            document.write('<param name="DataParam" value="' + dataUrlParams + '">');
        }
        document.write('<param name="SaveReportURL" value="' + saveurl + '">');
    }
    else {
        if (!ajaxMode) {
            document.write('<param name="ReportURL" value="' + report + '">');
            document.write('<param name="DataURL" value="' + data + '">');
            document.write('<param name="DataParam" value="' + dataUrlParams + '">');
        }
        document.write('<param name="AutoRun" value=' + _autorun + '>');
    }
    document.write('<param name="SerialNo" value="' + gr_SerialNo + '">');
    document.write('<param name="UserName" value="' + gr_UserName + '">');
    exparams && document.write(exparams);
    document.write('</object>');

    plugin = document.getElementById(id);

    //(!plugin || !plugin.Report) && MsgPluginFailed();
    if (!plugin || !plugin.Report) {
        MsgPluginFailed();
        return;
    }
    oncreate && oncreate(plugin);

    if (ajaxMode) {
        if (isdesigner) {
            if (report) {
                AjaxDesignerOpen(plugin, report, args.onreportload);
            }

            if (data && !plugin.OnRequestData) {
                plugin.OnRequestData = function (Report) {
                    SyncReportLoadData(Report, data); //这里必须是同步方式
                }
            }

            if (saveurl && !plugin.OnSaveReport) {
                plugin.OnSaveReport = function () {
                    AjaxDesignerSave(plugin, saveurl);
                    plugin.DefaultAction = false;
                    alert("报表模板设计已提交至服务器！\r\n(此提示信息可以在 CreateControl.js 文件中修改)");
                }
            }
        }
        else {
            _viewerStart(plugin, args);
        }
    }
}

//创建报表查询显示插件，详细请查看帮助中的 IGRDisplayViewer
function InsertDisplayViewer(args) {
    args.type = "displayviewer";
    doInsertPlugin(args);
}

//创建报表打印显示插件，详细请查看帮助中的 IGRPrintViewer
function InsertPrintViewer(args) {
    args.type = "printviewer";
    doInsertPlugin(args);
}

//创建报表设计器插件，详细请查看帮助中的 IGRDesigner
function InsertDesigner(args) {
    args.type = "designer";
    doInsertPlugin(args);
}

//////////////////////////////////////////////////////////////////////////
//按 AJAX 异步方式请求报表模板与数据，在数据都响应后执行报表插件方法
function _doAjaxReport(Report, ReportUrl, DataUrl, fun) {
    grplugin_http(ReportUrl, function (xmlhttp) {
        Report.LoadFromStr(xmlhttp.responseText);

        if (DataUrl) {
            grplugin_http(DataUrl, function (xmlhttp) {
                Report.LoadDataFromAjaxRequest(xmlhttp.responseText, xmlhttp.getAllResponseHeaders()); //加载报表数据

                fun();
            });
        }
        else {
            fun();
        }
    });
}

function AjaxReportPrint(Report, ReportUrl, DataUrl, ShowPrintDialog) {
    _doAjaxReport(Report, ReportUrl, DataUrl, function () {
        Report.Print(ShowPrintDialog);
    });
}

function AjaxReportPrintPreview(Report, ReportUrl, DataUrl, ShowModal) {
    _doAjaxReport(Report, ReportUrl, DataUrl, function () {
        Report.PrintPreview(ShowModal);
    });
}

function AjaxReportExportDirect(Report, ReportUrl, DataUrl, ExportType, FileName, ShowOptionDlg, DoneOpen) {
    _doAjaxReport(Report, ReportUrl, DataUrl, function () {
        Report.ExportDirect(ExportType, FileName, ShowOptionDlg, DoneOpen);
    });
}

function AjaxDesignerOpen(designer, LoadReportURL, onreportload) {
    grplugin_http(LoadReportURL, function (xmlhttp) {
        designer.Report.LoadFromStr(xmlhttp.responseText);
        designer.Reload();
        onreportload && onreportload(designer.Report);
    });
}

function AjaxDesignerSave(designer, SaveReportURL) {
    var str;

    designer.Post();
    str = designer.Report.SaveToStr();

    grplugin_http(SaveReportURL, 0, true, str);
}


//////////////////////////////////////////////////////////////////////////
//按同步方式请求报表数据，数据请求方法调用后紧接着调用报表载入数据的方法

//从URL处加载报表模板
function SyncReportLoad(Report, ReportUrl) {
    grplugin_http(ReportUrl, function (xmlhttp) {
        Report.LoadFromStr(xmlhttp.responseText);
    }, false);
}

//从URL处加载报表数据
function SyncReportLoadData(Report, DataUrl) {
    grplugin_http(DataUrl, function (xmlhttp) {
        Report.LoadDataFromAjaxRequest(xmlhttp.responseText, xmlhttp.getAllResponseHeaders()); //加载报表数据
    }, false);
}

//从URL处加载记录集数据
function SyncRecordsetData(Recordset, DataUrl) {
    grplugin_http(DataUrl, function (xmlhttp) {
        Recordset.LoadDataFromXML(xmlhttp.responseText);
    }, false);
}

//从URL处加载图表数据，参数说明见 IGRChart.LoadDataFromXML 方法
function SyncChartData(Chart, DataUrl, FirstSeries, AutoSeries, AutoGroup) {
    grplugin_http(DataUrl, function (xmlhttp) {
        Chart.LoadDataFromXML(xmlhttp.responseText, FirstSeries, AutoSeries, AutoGroup);
    }, false);
}

//从URL处加载XY类型的图表数据，参数说明见 IGRChart.LoadXYDataFromXML 方法
function SyncChartXYData(Chart, DataUrl, AutoSeries) {
    grplugin_http(DataUrl, function (xmlhttp) {
        Chart.LoadXYDataFromXML(xmlhttp.responseText, AutoSeries);
    }, false);
}

//从URL处加载XYZ类型的图表数据，参数说明见 IGRChart.LoadXYZDataFromXML 方法
function SyncChartXYZData(Chart, DataUrl, AutoSeries) {
    grplugin_http(DataUrl, function (xmlhttp) {
        Chart.LoadXYZDataFromXML(xmlhttp.responseText, AutoSeries);
    }, false);
}

//从URL处加载图像框的图像据，Url响应的数据必须是BASE64编码的文本数据
function SyncPictureBoxData(PictureBox, Url, Report) { //下一版本PictureBox可以取得Report，则最后一个参数可以不要
    grplugin_http(Url, function (xmlhttp) {
        var bin = Report.Utility.CreateBinaryObject();

        bin.LoadFromVariant(xmlhttp.responseText);
        PictureBox.LoadFromBinary(bin);
    }, false);
}

//////////////////////////////////////////////////////////////////////////
//以下函数是为了兼容以前的写法
function CreatePrintViewerEx2(PluginID, Width, Height, ReportURL, DataURL, AutoRun, ExParams) {
    doInsertPlugin({
        type: "printviewer",
        id: PluginID,
        width: Width,
        height: Height,
        report: ReportURL,
        data: DataURL,
        autorun: AutoRun,
        exparams: ExParams
    });
}

function CreateDisplayViewerEx2(PluginID, Width, Height, ReportURL, DataURL, AutoRun, ExParams, OnCreated, OnReportLoaded) {
    doInsertPlugin({
        type: "displayviewer",
        id: PluginID,
        width: Width,
        height: Height,
        report: ReportURL,
        data: DataURL,
        autorun: AutoRun,
        exparams: ExParams
    });
}

function CreateDesignerEx(Width, Height, LoadReportURL, SaveReportURL, DataURL, ExParams) {
    doInsertPlugin({
        type: "designer",
        width: Width,
        height: Height,
        report: LoadReportURL,
        data: DataURL,
        saveurl: SaveReportURL,
        exparams: ExParams
    });
}

function CreatePrintViewerEx(Width, Height, ReportURL, DataURL, AutoRun, ExParams) {
    CreatePrintViewerEx2("ReportViewer", Width, Height, ReportURL, DataURL, AutoRun, ExParams)
}

function CreateDisplayViewerEx(Width, Height, ReportURL, DataURL, AutoRun, ExParams) {
    CreateDisplayViewerEx2("ReportViewer", Width, Height, ReportURL, DataURL, AutoRun, ExParams)
}

function CreatePrintViewer(ReportURL, DataURL) {
    CreatePrintViewerEx("100%", "100%", ReportURL, DataURL, true, "");
}

function CreateDisplayViewer(ReportURL, DataURL) {
    CreateDisplayViewerEx("100%", "100%", ReportURL, DataURL, true, "");
}

function CreateDesigner(LoadReportURL, SaveReportURL, DataURL) {
    CreateDesignerEx("100%", "100%", LoadReportURL, SaveReportURL, DataURL, "");
}

//////////////////////////////////////////////////////////////////////////
//应用工具类函数

//HTTP通讯获取数据函数。参数async为true为异步方式，默认为异步。
function grplugin_http(url, callback, async, url_params, url_method, cbthis) {
    var xmlhttp = new XMLHttpRequest();

    function method_valid(url, method) {
        return method ? method : (/.grf|.txt|.xml|.json|.png|.jpg|.jpeg|.bmp|.gif/.test(url) ? "GET" : "POST");
    }

    xmlhttp.onreadystatechange = function () {
        if (xmlhttp.readyState == 4 && xmlhttp.status > 0) {
            if (xmlhttp.status == 200) {
                callback && callback.call(cbthis, xmlhttp);
            }
            else {
                window.open(url, "blank");
            }
        }
    }

    xmlhttp.onerror = function () {
        window.open(url, "blank");
    }

    //如果参数async没有定义，则默认为异步请求数据
    async = async || (async == undefined);

    xmlhttp.open(method_valid(url, url_method), url, async);
    xmlhttp.send(url_params);  //POST 或 PUT 可以传递参数
}

//此函数用于判断一个变量是否为URL字符串，如果类型为字符串且首个非空白字符不为“<”与“{”即判定为URL
var grplugin_is_url = function (p) {
    var index = 0,
        len = p.length,
        ch;

    if (typeof p != "string") {
        return 0;
    }

    //首先找到第一个非空白字符
    while (index < len) {
        ch = p[index];
        if (!/\s/g.test(ch))
            break;
        index++;
    }

    return (ch != "{") && (ch != "<") && (p.substr(index, 4) != "_WR_");
};

//为URL追加一个名为id的随机数参数，用于防止浏览器缓存。
//报表模板重新设计后，因为浏览器缓存而让报表生成不能反映出新修改的设计结果，URL后追加一个随机数参数可以避免这样的问题
//参数url必须是静态的URL，其后本身无任何参数
//如果模板几乎不怎么修改，可以去掉对本函数的调用
function urlAddRandomNo(url) {
    return url + "?id=" + Math.floor(Math.random() * 10000);
}

//由三原色值合成颜色整数值
function ColorFromRGB(red, green, blue) {
    return red + green * 256 + blue * 256 * 256;
}

//获取颜色中的红色值，传入参数为整数表示的RGB值
function ColorGetR(intColor) {
    return intColor & 255;
}

//获取颜色中的绿色值，传入参数为整数表示的RGB值
function ColorGetG(intColor) {
    return (intColor >> 8) & 255;
}

//获取颜色中的蓝色值，传入参数为整数表示的RGB值
function ColorGetB(intColor) {
    return (intColor >> 16) & 255;
}