/// <reference path="../App.js" />
var username;
var password;
var username_id;
var key;

var selectedIndex;

var listSSSElement;
var listView;
var listViewOnExcel;

var sssServer = "https://www.spreadsheetspace.net";

//var worker;
var overView = false;

var coeff, d, dmp1, dmq1, e, n, p, q;

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $("#server").val(sssServer);
            $("#content-sss").hide();

            var u = "";
            var p = "";
            var s = "";
            //var k = "";
            var u_id = "";

            listViewOnExcel = [];

            u = getCookie("username");
            p = getCookie("password");
            //s = getCookie("server");
            u_id = getCookie("username_id");
            //k = getCookie("rsa");
            var logged = false;

            //if (u != "" && p != "" && u_id != "" && s != "") {
            if (u != "" && p != "" && u_id != "") {
                logged = true;
            }
            //controllo se mi sono gia' loggato e ho le credenziali salvate
            if (logged) {
                username = u;
                password = p;
                //sssServer = s;
                username_id = u_id;

                //faccio il login automatico
                //$("#server").val(sssServer);
                $("#username").val(username);
                $("#password").val(password);

                login();

                /*setCookie("rsa", "", -1);
                k = getCookie("rsa");
                if (k == "") {
                    //generateRsaKey();
                } else {
                    //setRSA("", false);
                }*/

            } else {
                //permetto all'utente di fare il login e/o registrarsi
                $('#logout').prop("disabled", true);

                $('#connectTab').prop("disabled", true);
                $('#exposeTab').prop("disabled", true);

                $("#content-sss").hide();
                $("#login-div").show();

                $("#welcome").text(username);

                $("#operationProgress").hide();

                listSSSElement = "";
            }

            document.getElementById("exposeTab").style.borderRight = "none";
            document.getElementById("homeTabSSS").style.display = "block";

            $('#homeTab').click(clickTab);
            $('#connectTab').click(clickTab);
            $('#exposeTab').click(clickTab);

            $('#login').click(login);
            $('#logout').click(logout);

            $('#refresh-view').click(getViews);
            $('#create-view').click(createView);

            $('#consoleWeb').click(consoleWeb);
        });
    };

    function clickTab(evt) {
        // Declare all variables
        var tabName = this.id;
        var i, tabcontent, tablinks;

        // Get all elements with class="tabcontent" and hide them
        tabcontent = document.getElementsByClassName("tabcontent");
        for (i = 0; i < tabcontent.length; i++) {
            tabcontent[i].style.display = "none";
        }

        // Get all elements with class="tablinks" and remove the class "active"
        tablinks = document.getElementsByClassName("tablinks");
        for (i = 0; i < tablinks.length; i++) {
            tablinks[i].className = tablinks[i].className.replace(" active", "");
        }

        // Show the current tab, and add an "active" class to the button that opened the tab
        document.getElementById(tabName + "SSS").style.display = "block";
        evt.currentTarget.className += " active";
    }

    function setCookie(cname, cvalue, exdays) {
        var d = new Date();
        d.setTime(d.getTime() + (exdays * 24 * 60 * 60 * 1000));
        var expires = "expires=" + d.toUTCString();
        document.cookie = cname + "=" + cvalue + ";" + expires + ";path=/";
    }

    function getCookie(cname) {
        var name = cname + "=";
        var ca = document.cookie.split(';');
        for (var i = 0; i < ca.length; i++) {
            var c = ca[i];
            while (c.charAt(0) == ' ') {
                c = c.substring(1);
            }
            if (c.indexOf(name) == 0) {
                return c.substring(name.length, c.length);
            }
        }
        return "";
    }

    
    /*function setRSA(rsa, isGenerated) {
        if (isGenerated) {
            coeff = b2s(rsa.coeff);
            d = b2s(rsa.d);
            dmp1 = b2s(rsa.dmp1);
            dmq1 = b2s(rsa.dmq1);
            e = b2s(rsa.e);
            n = b2s(rsa.n);
            p = b2s(rsa.p);
            q = b2s(rsa.q);
        } else {
            coeff = getCookie("rsa_coeff");
            d = getCookie("rsa_d");
            dmp1 = getCookie("rsa_dmp1");
            dmq1 = getCookie("rsa_dmq1");
            e = getCookie("rsa_e");
            n = getCookie("rsa_n");
            p = getCookie("rsa_p");
            q = getCookie("rsa_q"); 
        }

        setCookie("rsa", "true", 365);
        setCookie("rsa_coeff", coeff, 365);
        setCookie("rsa_d", d, 365);
        setCookie("rsa_dmp1", dmp1, 365);
        setCookie("rsa_dmq1", dmq1, 365);
        setCookie("rsa_e", e, 365);
        setCookie("rsa_n", n, 365);
        setCookie("rsa_p", p, 365);
        setCookie("rsa_q", q, 365);
    }*/

    function b2s(array) {
        var result = "";

        for (var i = 0; i < array.t; i++) {
            result += (String.fromCharCode(array[i]));
        }

        return result;
    }

    function resetCookie() {
        setCookie("username", "", -1);
        setCookie("password", "", -1);
        setCookie("username_id", "", -1);

        /*setCookie("rsa", "", -1);
        setCookie("rsa_coeff", "", -1);
        setCookie("rsa_d", "", -1);
        setCookie("rsa_dmp1", "", -1);
        setCookie("rsa_dmq1", "", -1);
        setCookie("rsa_e", "", -1);
        setCookie("rsa_n", "", -1);
        setCookie("rsa_p", "", -1);
        setCookie("rsa_q", "", -1);*/
    }

    function generateUUID() {
        var d = new Date().getTime();
        if (typeof performance !== 'undefined' && typeof performance.now === 'function') {
            d += performance.now(); //use high-precision timer if available
        }
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
            var r = (d + Math.random() * 16) % 16 | 0;
            d = Math.floor(d / 16);
            return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
        });
    }

    /*function generateRsaKey() {
        var rsa;
        if (typeof (Worker) !== "undefined") {
            worker = new Worker("generate_rsa.js");
            worker.postMessage(username + username_id);
            worker.onmessage = function (event) {
                rsa = event.data;

                setRSA(rsa, true);
                //setCookie("rsa", "true", 365);
            }
        }
    }*/

    function login() {
        sssServer = $("#server").val();
        username = $("#username").val();
        password = $("#password").val();
        addSelectionChangedEventHandler();

        $("#operationProgress").show();

        var url = sssServer + '/orchestrator/login?action=loginFromAddin';

        $.ajax({
            url: url,
            type: 'POST',
            data: null,
            headers: { 'X-Username': username, 'X-Password': password },
            success: function (data, textStatus, jqXHR) {
                if (data.statusCode == 200) {
                    app.showNotification('Logged-in');

                    setCookie("server", sssServer, 365);
                    setCookie("username", username, 365);
                    setCookie("password", password, 365);
                    setCookie("username_id", data.userId, 365);

                    username_id = data.userId;

                    $('#logout').prop("disabled", false);

                    $('#connectTab').prop("disabled", false);
                    $('#exposeTab').prop("disabled", false);

                    $("#content-sss").show();
                    $("#login-div").hide();

                    $("#welcome").text(username);

                    getViews("", -1, "", "", "", "", "");
                    $("#operationProgress").hide();
                } else {
                    app.showNotification('Error on login (1)');
                    $("#operationProgress").hide();
                }
            },
            error: function (jqXHR, textStatus, errorThrown) {
                app.showNotification('Error on login (2)');
                $("#operationProgress").hide();

            }
        });
    }

    function logout() {
        //svuoto lo storage salvato
        resetCookie();

        listSSSElement = "";
        listView = [];

        //faccio il reset delle caselle di testo
        //$("#server").val("");
        //$("#username").val("");
        $("#password").val("");

        //mostro il div per il login
        $("#login-div").show();

        //nascondo la possibilità di fare le operazioni su sss
        $("#content-sss").hide();

        $("#welcome").text('');

        $("#operationProgress").hide();

        //disabilito il bottone di logout
        $('#logout').prop("disabled", true);

        $('#connectTab').prop("disabled", true);
        $('#exposeTab').prop("disabled", true);

        var viewElement;
        for (var i = 0; i < listViewOnExcel.length; i++) {
            viewElement = listViewOnExcel.splice(i, 1);
            checkBindingToRemove(viewElement[0].local_uuid);
        }

        tableView();
    }

    function consoleWeb() {
        username = $("#username").val();
        password = $("#password").val();

        /*var console = sssSite;
        console = console.split("$USERNAME$").join(username);
        console = console.split("$PASSWORD$").join(password);
        console = console.split("$SHOW_ADDRESS_BOOK$").join(false);
        console = console.split("$SERVER_URL$").join(sssServer + "/orchestrator/login");

        var newWindow = window.open();
        newWindow.document.write(console);*/
    }



    function addSelectionChangedEventHandler() {
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, MyHandler);
        //Office.context.document.addHandlerAsync(Office.EventType.BindingDataChanged, MyHandler2);
        //Office.context.document.addHandlerAsync(Office.EventType.ViewSelectionChanged, MyHandler);
        Office.context.document.addHandlerAsync(Office.EventType.ResourceSelectionChanged, MyHandler);
    }


    function MyHandler(eventArgs) {
               
        eventArgs.document.getSelectedDataAsync(Office.CoercionType.Matrix,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    Excel.run(function (ctx) {
                        var selectedRange = ctx.workbook.getSelectedRange();
                        selectedRange.load('address');

                        return ctx.sync().then(function () {
                            document.getElementById('text').innerText = eventArgs.type + " " + result.value + " " + selectedRange.address;
                        });
                    }).catch(function (error) {
                        console.log("Error: " + error);
                    });
                }
            }
        );
    }

    function addNewBinding(rangeAddress, local_uuid) {
        /*Office.context.document.bindings.addFromNamedItemAsync(rangeAddress, "matrix", {id: local_uuid}, 
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    console.log("error");
                } else {
                    console.log("ok");
                }
            }
        );*/

        Office.context.document.bindings.addFromNamedItemAsync(rangeAddress, "matrix", { id: local_uuid });
        Office.select("bindings#" + local_uuid).addHandlerAsync(Office.EventType.BindingDataChanged, 
            function (eventArgs) {
                checkAddRowOrColumn(eventArgs.binding.id);
            }
        );
    }

    function checkAddRowOrColumn(local_uuid) {
        var element;
        for(var i = 0; i < listViewOnExcel.length; i++) {
            var el = listViewOnExcel[i];
            var el_uuid = el.local_uuid;

            if (local_uuid == el_uuid) {
                element = el;
                break;
            }
        }
        Excel.run(function (ctx) {
            var binding = ctx.workbook.bindings.getItem(local_uuid);
            var range = binding.getRange();
            range.load('address');

            return ctx.sync().then(function () {
                var rangeAddress = range.address;

                if (rangeAddress != element.rangeAddress) {
                    element.rangeAddress = rangeAddress;
                    tableView();
                }
            });
        }).catch(function (error) {
            console.log("Error: " + error);
        });
    }

    function checkBindingToRemove(local_uuid) {
        Office.select("bindings#" + local_uuid).removeHandlerAsync(Office.EventType.BindingDataChanged,
            function (eventArgs) {
                checkAddRowOrColumn(eventArgs.binding.id);
            }
        );

        Office.context.document.bindings.releaseByIdAsync(local_uuid, 
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    console.log("error");
                } else {
                    console.log("ok");
                }
            }
        );
    }

    function addProtection(rangeAddress, permissionUser) {
        var index = rangeAddress.indexOf("!") + 1;
        var sheetName = rangeAddress.substring(0, index - 1);
        var rangeName = rangeAddress.substring(index);

        if (permissionUser == "WRITE") {
            Excel.run(function (ctx) {
                var sheet = ctx.workbook.worksheets.getItem(sheetName);
                var range = sheet.getRange(rangeName);
                range.load(["format/", "format/protection/", "format/protection/locked", "address"]);
                
                return ctx.sync().then(function () {
                    var rA = range.address;
                    range.format.protection.locked = true;
                    sheet.protection.protect({ allowInsertRows: false });
                });
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    }
    
    function addItemOnTable(rangeAddress, seqNum, fileList, optionalFileList, usersPermission, indexListView) {
        var element;
        if (indexListView == undefined) {
            element = listView[selectedIndex - 1];
        } else {
            element = listView[indexListView - 1];
        }
        var description = element.description;
        var ownerMail = element.owner.mail;
        var owner = element.owner;
        var uuid = element.id;
        var viewServer = element.view_server;
        var category = element.category;
        var usersPermission = element.usersPermissionObj;

        var metadata = JSON.parse(element.metadata);
        var cols = metadata.cols;
        var rows = metadata.rows;
        var is_table = metadata.is_table;
        var has_headers = metadata.has_headers;
        var excelType = metadata.excelType;

        var permissionUser;
        for (var j = 0; j < element.usersPermissionObj.length; j++) {
            var uP = element.usersPermissionObj[j];
            if (uP.userPermission.mail == username) {
                if (uP.userPermission.permissionType == 0) {
                    permissionUser = "READ";
                } else {
                    permissionUser = "WRITE";
                }
                break;
            }
        }

        var local_uuid = generateUUID();

        listViewOnExcel.push({
            description: description,
            ownerMail: ownerMail,
            owner: owner,
            uuid: uuid,
            local_uuid: local_uuid,
            viewServer: viewServer,
            permissionUser: permissionUser,
            seqNum: seqNum,
            category: category,
            cols: cols,
            rows: rows,
            fileList: fileList,
            optionalFileList: optionalFileList,
            is_table: is_table,
            excelType: excelType,
            has_headers: has_headers,
            usersPermission: usersPermission,
            rangeAddress: rangeAddress
        });

        addNewBinding(rangeAddress, local_uuid);
        
        tableView();

        addProtection(rangeAddress, permissionUser);
    }

    function tableView() {
        if (listViewOnExcel.length > 0) {
            var table = '<table id="tableView" style="border-collapse: collapse;">\n<tbody>\n';
            for (var i = 0; i < listViewOnExcel.length; i++) {
                var listViewOnExcel_i = listViewOnExcel[i];
                table += '<tr id="tr">\n<td align="center"><b>' + listViewOnExcel_i.description + '</b></td>\n';
                table += '<td align="center">' + listViewOnExcel_i.ownerMail + " - (" + listViewOnExcel_i.permissionUser + ")" + '</td>\n';
                table += '<td align="center">' + listViewOnExcel_i.rangeAddress + '</td></tr>\n';
                table += '<tr id="tr" class="border_bottom">\n<td align="center">';
                if (listViewOnExcel_i.permission == "READ") {
                    table += '<button type="button" style="cursor: pointer;" id="update_' + i + '" title="Update" disabled><span class="fas fa-share-square"></span></button>' + '</td>\n';
                } else {
                    table += '<button type="button" style="cursor: pointer;" id="update_' + i + '" title="Update"><span class="fas fa-share-square"></span></button>' + '</td>\n';
                }
                table += '<td align="center">' + '<button type="button" style="cursor: pointer;" id="refresh_' + i + '" title="Refresh"><span class="fas fa-sync-alt"></span></button>' + '</td>\n';
                table += '<td align="center">' + '<button type="button" style="cursor: pointer;" id="remove_' + i + '" title="Unlink"><span class="far fa-times-circle"></span></button>' + '</td>\n</tr>\n';

            }
            table += '</tbody>\n</table>';
            document.getElementById("div-table-view").innerHTML = table;

            for (var i = 0; i < listViewOnExcel.length; i++) {
                var listViewOnExcel_i = listViewOnExcel[i];
                $('#update_' + i).on('click', { item: listViewOnExcel_i, index: i }, updateView);
                $('#refresh_' + i).on('click', { item: listViewOnExcel_i, index: i }, refreshView);
                $('#remove_' + i).on('click', { index: i }, removeView);
            }
        } else {
            var table = "";
            document.getElementById("div-table-view").innerHTML = table;
        }
    }
    
    function createView() {
        var resultTable;
        var recipient = $("#input-recipients").val();
        var description = $("#input-description").val();

        $("#operationProgress").show();

        var regexMail = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;

        var proceed = true;
        if (recipients == "" || description == "") {
            app.showNotification('Set recipients and/or description data');
            proceed = false;
        }

        var recipients = recipient.split(";");
        for (var i = 0; i < recipients.length; i++) {
            var r = recipients[i].trim();

            if(!regexMail.test(r.toLowerCase())) {
                app.showNotification('Set valid mail address');
                proceed = false;
            }
        }


        if (proceed) {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Table,
                 function (result) {
                     resultTable = result.status;
                     if (result.status === Office.AsyncResultStatus.Succeeded) {

                         var value = "";
                         var data = result.value;
                         var headers = data.headers;
                         var rowsData = data.rows;
                         var row = rowsData.length;
                         var column;

                         var is_table = true;
                         var has_headers;
                         if (headers.length == 1) {
                             has_headers = true;
                             for (var i = 0; i < headers.length; i++) {
                                 var cell = headers[i];

                                 column = cell.length;
                                 for (var j = 0; j < cell.length; j++) {
                                     value += cell[j];
                                     if (j != cell.length - 1) {
                                         value += "\x1F";
                                     }
                                 }

                                 if (i != headers.length - 1) {
                                     value += '\x1E';
                                 }
                             }
                         } else {
                             has_headers = false;
                         }
                         value += '\x1E';
                         for (var i = 0; i < rowsData.length; i++) {
                             var cell = rowsData[i];

                             column = cell.length;
                             for (var j = 0; j < cell.length; j++) {
                                 value += cell[j];
                                 if (j != cell.length - 1) {
                                     value += "\x1F";
                                 }
                             }

                             if (i != rowsData.length - 1) {
                                 value += '\x1E';
                             }
                         }

                         Excel.run(function (ctx) {
                             var selectedRange = ctx.workbook.getSelectedRange();
                             selectedRange.load('address');
                             return ctx.sync().then(function () {
                                 console.log(selectedRange.address);
                                 retrieveViewServer(value, row, column, is_table, has_headers, "VIEW", selectedRange.address, selectedRange);
                             });
                         }).catch(function (error) {
                             $("#operationProgress").hide();
                             console.log("Error: " + error);
                         });
                     }
                 }
            );

            if (resultTable == "failed" || resultTable == undefined) {
                Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,
                     function (result) {
                         if (resultTable == "failed" || resultTable == undefined) {
                             if (result.status === Office.AsyncResultStatus.Succeeded) {

                                 var value = "";
                                 var data = result.value;
                                 var row = data.length;
                                 var column;
                                 for (var i = 0; i < data.length; i++) {
                                     var cell = data[i];

                                     column = cell.length;
                                     for (var j = 0; j < cell.length; j++) {
                                         value += cell[j];
                                         if (j != cell.length - 1) {
                                             value += "\x1F";
                                         }
                                     }

                                     if (i != data.length - 1) {
                                         value += '\x1E';
                                     }
                                 }
                                 if (resultTable == "failed" || resultTable == undefined) {
                                     Excel.run(function (ctx) {
                                         var selectedRange = ctx.workbook.getSelectedRange();
                                         selectedRange.load('address');
                                         return ctx.sync().then(function () {
                                             console.log(selectedRange.address);
                                             if (resultTable == "failed" || resultTable == undefined) {
                                                 retrieveViewServer(value, row, column, false, false, "VIEW", selectedRange.address, selectedRange);
                                             }
                                         });
                                     }).catch(function (error) {
                                         $("#operationProgress").hide();
                                         console.log("Error: " + error);
                                     });
                                 }


                             } else {
                                 $("#operationProgress").hide();
                                 app.showNotification('Error:', result.error.message);
                             }
                         }
                     }
                 );
            }

        }
    }

    function getViews(rangeAddress, seqNum, fileList, optionalFileList, usersPermission, view_id, view_server) {
        sssServer = $("#server").val();
        username = $("#username").val();
        password = $("#password").val();
        listView = [];

        var url = sssServer + '/orchestrator/viewsMgmt/getViews';

        if (seqNum == undefined || fileList == undefined || optionalFileList == undefined || usersPermission == undefined || view_id == undefined || view_server == undefined) {
            rangeAddress = "";
            seqNum = -1;
            fileList = "";
            optionalFileList = "";
            usersPermission = "";
            view_id = "";
            view_server = "";
        }

        $.ajax({
            url: url,
            type: 'GET',
            data: null,
            cache: false,
            headers: { 'X-Username': username, 'X-Password': password },
            success: function (data, textStatus, jqXHR) {
                listSSSElement = data;

                createViewSelectHtml(data);

                if (view_id != "" && view_server != "") {
                    var indexListView = -1;
                    for (var i = 0; i < data.length; i++) {
                        var el = data[i];
                        if (el.id == view_id && el.view_server == view_server) {
                            indexListView = i + 1;
                            break;
                        }
                    }
                    addItemOnTable(rangeAddress, seqNum, fileList, optionalFileList, usersPermission, indexListView);
                }

                $("#operationProgress").hide();
            },
            error: function (jqXHR, textStatus, errorThrown) {
                app.showNotification('Error on getViews (1)');
                document.getElementById("div-list-view").innerHTML = "";
                $("#operationProgress").hide();

            }
        });
    }

    function createViewSelectHtml(data) {
        var select = "<select id=\"select-list-view\">\n"
        select += "<option>" + "Select View..." + "</option>\n";
        for (var i = 0; i < data.length; i++) {
            var item = data[i];
            var metadata = JSON.parse(item.metadata);

            if (item.encrypted == false) {
                if (metadata.excelType == "VIEW") {
                    if (metadata.cols != -1 && metadata.rows != -1) {
                        var description = item.description;
                        var owner = item.owner.mail;

                        listView.push(item);

                        select += "<option>" + description + " (" + owner + ")" + "</option>\n";
                    }
                }
            }
        }
        select += "</select>\n"
        document.getElementById("div-list-view").innerHTML = select;
        document.getElementById("select-list-view").onchange = function () {
            selectedIndex = this.selectedIndex;
            createViewSelectHtml(data);

            var view = listView[selectedIndex - 1];

            var permissionUser;
            for (var j = 0; j < view.usersPermissionObj.length; j++) {
                var uP = view.usersPermissionObj[j];
                if (uP.userPermission.mail == username) {
                    if (uP.userPermission.permissionType == 0) {
                        permissionUser = "READ";
                    } else {
                        permissionUser = "WRITE";
                    }
                    break;
                }
            }

            Excel.run(function (ctx) {
                var selectedRange = ctx.workbook.getSelectedRange();
                selectedRange.load('address');

                return ctx.sync().then(function () {
                    //creo il range con i dati calcolati prima ed incollo il risultato ottenuto dal ricalcolo
                    overView = false;
                    retrieveIntersection(selectedRange.address, view, 0);
                    //pullView(view);

                }).catch(function (error) {
                    console.log('Error:', error);
                });
            }).catch(function (error) {
                console.log("Error: " + error);
            });
        }
    }

    function retrieveIntersection(selectedRangeAddress, view, k) {
        var index = selectedRangeAddress.indexOf("!") + 1;
        var sheetName = selectedRangeAddress.substring(0, index - 1);

        if (listViewOnExcel.length == 0) {
            pullView(view);
        } else {
            Excel.run(function (ctx) {
                var view_k = listViewOnExcel[k];

                var rangeAddress_i = view_k.rangeAddress;
                var permission = view_k.permissionUser;

                index = rangeAddress_i.indexOf("!") + 1;
                var sheetName_i = rangeAddress_i.substring(0, index - 1);

                if (sheetName == sheetName_i) {
                    var selectedRange = ctx.workbook.worksheets.getItem(sheetName).getRange(selectedRangeAddress);

                    var rangeIntersection = selectedRange.getIntersection(rangeAddress_i);
                    rangeIntersection.load('address');
                }

                return ctx.sync().then(function () {
                    if (!overView) {
                        var addressIntersection = rangeIntersection.address;
                        k = listViewOnExcel.length - 1;

                        if (permission == "READ") {
                            overView = true;
                        }
                    }

                    if (k == listViewOnExcel.length - 1) {
                        if (overView) {
                            app.showNotification("Can't manage new view over another one");
                        } else {
                            pullView(view);
                        }
                    } else {
                        k = k + 1;
                        retrieveIntersection(selectedRangeAddress, view, k);
                    }
                });
            }).catch(function (error) {
                if (k == listViewOnExcel.length - 1) {
                    if (overView) {
                        app.showNotification("Can't manage new view over another one");
                    } else {
                        pullView(view);
                    }
                } else {
                    k = k + 1;
                    retrieveIntersection(selectedRangeAddress, view, k);
                }
            });
        }
    }

    function retrieveViewServer(value, row, column, is_table, has_headers, type, selectedRangeAddress, range) {
        sssServer = $("#server").val();
        username = $("#username").val();
        password = $("#password").val();

        var status = 400;

        var url = sssServer + "/orchestrator/viewserver";

        var hostname;

        $.ajax({
            url: url,
            type: 'GET',
            data: null,
            headers: { 'X-Username': username, 'X-Password': password },
            contentType: 'application/json;charset=utf-8',
            success: function (data, status, jqXHR) {
                hostname = data.hostname;

                pushView(value, hostname, row, column, is_table, has_headers, type, selectedRangeAddress, range);
            },
            error: function (jqXHR, status, errorThrown) {
                app.showNotification("Error during retrieveViewserver (1)");
                $("#operationProgress").hide();
            }

        });
    }

    function pushView(value, viewserver, row, column, table, header, type, selectedRangeAddress, range) {
        sssServer = $("#server").val();
        username = $("#username").val();
        password = $("#password").val();

        var recipients = $("#input-recipients").val();
        var description = $("#input-description").val();

        var sender, recipient, owner;
        var senderEncrypted, recipientEncrypted;
        var id_sender, id_recipient;

        var objPushView;
        sender = username;
        recipient = recipients.split(";");
        var usersPermission = [];

        var mail, key;
        senderEncrypted = false;
        recipientEncrypted = false;

        var encrypted;
        if (senderEncrypted && recipientEncrypted) {
            encrypted = true;
        } else {
            encrypted = false;
        }

        owner = username_id;
        id_sender = id_recipient = "";

        var description = description;
        var viewServer = viewserver;
        var fromAddin = true;

        var is_table = table;;
        var has_headers = header;;
        var rows = row;
        var cols = column;
        var has_template = false;
        var has_data = true;
        var excelType = type;

        var tmpMetadata = new Metadata(is_table, has_headers, rows, cols, has_template, has_data, excelType);
        var metadata = JSON.stringify(tmpMetadata);

        var userPermission = new UsersPermissionWs(sender, 1);
        usersPermission.push(userPermission);
        for (var i = 0; i < recipient.length; i++) {
            var r = recipient[i].trim();

            userPermission = new UsersPermissionWs(r, 0);
            usersPermission.push(userPermission);
        }

        var id = "";
        var clientType = 0;

        var filename, sequenceNumber, type, data;
        var Encrypt, symmetricKeys, dataData;
        var randomKey;
        var str64;

        if (encrypted) {
            Encrypt = true;
            //randomKey = RandomKey();
            //symmetricKeys = encryptSymmetric(randomKey, id_sender, id_recipient);
        } else {
            Encrypt = false;
            symmetricKeys = null;
        }

        var optionalFileList = [];
        var fileList = [];

        /*filename = "template";
        sequenceNumber = 0;
        type = "OPTIONAL";
        var byteData = []; //FIXME

        if (encrypted) {
            str64 = window.btoa(byteData);
            dataData = encryptDES(str64, randomKey);
        } else {
            str64 = window.btoa(byteData);
            dataData = str64;
        }


        var tmpData = new Data(Encrypt, symmetricKeys, dataData);
        var strData = JSON.stringify(tmpData);
        byteData = [];
        for (var i = 0; i < strData.length; ++i) {
            byteData.push(strData.charCodeAt(i));
        }
        var data = btoa(String.fromCharCode.apply(null, new Uint8Array(byteData)));

        var optionalFile = new FileList(filename, sequenceNumber, type, data);
        optionalFileList.push(optionalFile);*/

        filename = null;
        sequenceNumber = 1;
        type = "FULL";
        var byteData = value;
        if (encrypted) {
            str64 = window.btoa(byteData);
            dataData = encryptDES(str64, randomKey);
        } else {
            str64 = window.btoa(byteData);
            dataData = str64;
        }
        var tmpData = new Data(Encrypt, symmetricKeys, dataData);
        var strData = JSON.stringify(tmpData);
        byteData = [];
        for (var i = 0; i < strData.length; ++i) {
            byteData.push(strData.charCodeAt(i));
        }
        data = btoa(String.fromCharCode.apply(null, new Uint8Array(byteData)));

        var file = new FileList(filename, sequenceNumber, type, data);
        fileList.push(file);

        objPushView = new PushRequest(id, description, owner, usersPermission, viewServer, fromAddin, optionalFileList, fileList, clientType, metadata);

        var jsonObjPushView = JSON.stringify(objPushView);

        var status = 400;

        var url = viewServer + "/viewserver/push";

        $.ajax({
            url: url,
            type: 'POST',
            data: jsonObjPushView,
            cache: false,
            headers: { 'X-Username': username, 'X-Password': password },
            contentType: 'application/json;charset=utf-8',
            success: function (data, status, jqXHR) {
                var result = JSON.parse(data);
                var statusCode = result.statusCode;
                if (statusCode == 200) {
                    var viewserverid = result.viewId;
                    eventCreation(viewserverid, description, owner, usersPermission, viewServer, fromAddin, clientType, metadata, encrypted, optionalFileList, fileList, selectedRangeAddress, range)
                } else {
                    if (statusCode == 408) {
                        app.showNotification(result.authorizationResponse.message);
                    } else {
                        app.showNotification('Error on pushView (1)');
                    }
                    $("#operationProgress").hide();
                }
            },
            error: function (jqXHR, status, errorThrown) {
                app.showNotification('Error on pushView (2)');
                $("#operationProgress").hide();
            }

        });
    }

    function eventCreation(id, description, owner, usersPermission, viewServer, fromAddin, clientType, metadata, encrypted, optionalFileList, fileList, selectedRangeAddress, range) {
        sssServer = $("#server").val();
        username = $("#username").val();
        password = $("#password").val();

        var objEventCreation = new ObjEventCreation(id, description, owner, usersPermission, viewServer, fromAddin, clientType, metadata, encrypted);

        var jsonObjEventCreation = JSON.stringify(objEventCreation);

        var status = 400;

        var url = sssServer + "/orchestrator/event/creation";

        $.ajax({
            url: url,
            type: 'POST',
            cache: false,
            data: jsonObjEventCreation,
            headers: { 'X-Username': username, 'X-Password': password },
            contentType: 'application/json;charset=utf-8',
            success: function (data, status, jqXHR) {
                app.showNotification('View successfully created');

                getViews(selectedRangeAddress, 0, fileList, optionalFileList, usersPermission, id, viewServer);

                //faccio il reset delle caselle di testo
                $("#input-recipients").val("");
                $("#input-description").val("");
            },
            error: function (jqXHR, status, errorThrown) {
                $("#operationProgress").hide();
                app.showNotification('Error on eventCreation');
            }

        });
    }

    function updateView(event) {
        var item = event.data.item;
        sssServer = $("#server").val();
        username = $("#username").val();
        password = $("#password").val();

        $("#operationProgress").show();

        Excel.run(function (ctx) {
            var index = item.rangeAddress.indexOf("!") + 1;
            var sheetName = item.rangeAddress.substring(0, index - 1);
            var rangeAddress = item.rangeAddress.substring(index);

            var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
            range.load('values');
            return ctx.sync().then(function () {
                var value = "";
                var data = range.values;
                var row = data.length;
                var column;
                for (var i = 0; i < data.length; i++) {
                    var cell = data[i];

                    column = cell.length;
                    for (var j = 0; j < cell.length; j++) {
                        value += cell[j];
                        if (j != cell.length - 1) {
                            value += "\x1F";
                        }
                    }

                    if (i != data.length - 1) {
                        value += '\x1E';
                    }
                }

                var optionalFileList = [];
                var fileList = [];

                var id = "";
                var clientType = 0;

                var filename, sequenceNumber, type, data;
                var Encrypt, symmetricKeys, dataData;
                var randomKey;
                var str64;

                var mail, key;
                var senderEncrypted = false;
                var recipientEncrypted = false;

                var usersPermission = [];
                var userPermission;
                for (var i = 0; i < item.usersPermission.length; i++) {
                    var uP = item.usersPermission[i];
                    userPermission = new UsersPermissionWs(uP.userPermission.mail, uP.userPermission.permissionType);
                    usersPermission.push(userPermission);
                }

                var encrypted;
                if (senderEncrypted && recipientEncrypted) {
                    encrypted = true;
                } else {
                    encrypted = false;
                }

                if (encrypted) {
                    Encrypt = true;
                    //randomKey = RandomKey();
                    //symmetricKeys = encryptSymmetric(randomKey, id_sender, id_recipient);
                } else {
                    Encrypt = false;
                    symmetricKeys = null;
                }

                /*filename = "template";
                sequenceNumber = 0;
                type = "OPTIONAL";
                var byteData = []; //FIXME
        
                if (encrypted) {
                    str64 = window.btoa(byteData);
                    dataData = encryptDES(str64, randomKey);
                } else {
                    str64 = window.btoa(byteData);
                    dataData = str64;
                }
        
        
                var tmpData = new Data(Encrypt, symmetricKeys, dataData);
                var strData = JSON.stringify(tmpData);
                byteData = [];
                for (var i = 0; i < strData.length; ++i) {
                    byteData.push(strData.charCodeAt(i));
                }
                var data = btoa(String.fromCharCode.apply(null, new Uint8Array(byteData)));
        
                var optionalFile = new FileList(filename, sequenceNumber, type, data);
                optionalFileList.push(optionalFile);*/

                filename = null;
                sequenceNumber = item.seqNum + 1;
                type = "FULL";
                var byteData = value;
                if (encrypted) {
                    str64 = window.btoa(byteData);
                    dataData = encryptDES(str64, randomKey);
                } else {
                    str64 = window.btoa(byteData);
                    dataData = str64;
                }
                var tmpData = new Data(Encrypt, symmetricKeys, dataData);
                var strData = JSON.stringify(tmpData);
                byteData = [];
                for (var i = 0; i < strData.length; ++i) {
                    byteData.push(strData.charCodeAt(i));
                }
                data = btoa(String.fromCharCode.apply(null, new Uint8Array(byteData)));

                var file = new FileList(filename, sequenceNumber, type, data);
                fileList.push(file);

                var objUpdateRequest = new PushRequest(item.uuid, "", item.owner.mail, usersPermission, item.viewServer, "", optionalFileList, fileList, "", "")
                var jsonObjUpdateRequest = JSON.stringify(objUpdateRequest);

                var status = 400;

                var url = item.viewServer + "/viewserver/push/" + item.uuid;

                $.ajax({
                    url: url,
                    type: 'POST',
                    cache: false,
                    data: jsonObjUpdateRequest,
                    headers: { 'X-Username': username, 'X-Password': password },
                    contentType: 'application/json;charset=utf-8',
                    success: function (data, status, jqXHR) {
                        var parsedData = JSON.parse(data);
                        var statusCode = parsedData.statusCode;
                        var nextNumber = parsedData.nextSeqNumAvailable;
                        if (statusCode != 200) {
                            if (statusCode == 403) {
                                app.showNotification("Can not send the update, the current version is not up to date with the latest version on the server SpreadSheetSpace. Your data may be overwritten.");
                            } else {
                                app.showNotification("Error during update view (1)");
                            }
                            $("#operationProgress").hide();
                        } else {
                            eventUpdate(item.uuid, item.viewServer, encrypted, item, event.data.index);
                        }
                    },
                    error: function (jqXHR, status, errorThrown) {
                        app.showNotification("Error during update view (2)");
                        $("#operationProgress").hide();
                    }

                });
            });
        }).catch(function (error) {
            app.showNotification("Error during update view (3)");
            $("#operationProgress").hide();
            if (error instanceof OfficeExtension.Error) {
                app.showNotification("Error during update view (4)");
            }
        });
    }

    function eventUpdate(id, viewServer, encrypted, item, index) {
        sssServer = $("#server").val();
        username = $("#username").val();
        password = $("#password").val();

        var objEventUpdate = new ObjEventUpdate(id, viewServer, encrypted);

        var jsonObjEventUpdate = JSON.stringify(objEventUpdate);

        var status = 400;

        var url = viewServer + "/orchestrator/event/update";

        $.ajax({
            url: url,
            type: 'POST',
            cache: false,
            data: jsonObjEventUpdate,
            headers: { 'X-Username': username, 'X-Password': password },
            contentType: 'application/json;charset=utf-8',
            success: function (data, status, jqXHR) {
                app.showNotification('View successfully updated');
                console.log(JSON.stringify(jqXHR));

                var seqNum = item.seqNum + 1;
                listViewOnExcel[index].seqNum = seqNum;
                $("#operationProgress").hide();
            },
            error: function (jqXHR, status, errorThrown) {
                app.showNotification('Error on eventUpdate (1)');
                $("#operationProgress").hide();
            }

        });
    }

    function pullView(element) {
        sssServer = $("#server").val();
        username = $("#username").val();
        password = $("#password").val();

        $("#operationProgress").show();

        var address, rowIndex, columnIndex;
        Excel.run(function (ctx) {
            var selectedRange = ctx.workbook.getSelectedRange();
            selectedRange.load('address');
            selectedRange.load('rowIndex');
            selectedRange.load('columnIndex');

            return ctx.sync().then(function () {
                //creo il range con i dati calcolati prima ed incollo il risultato ottenuto dal ricalcolo
                address = selectedRange.address;
                rowIndex = selectedRange.rowIndex;
                columnIndex = selectedRange.columnIndex;

            }).catch(function (error) {
                console.log('Error:', error);
                $("#operationProgress").hide();
            });
        }).catch(function (error) {
            console.log("Error: " + error);
            $("#operationProgress").hide();
        });

        var id = element.id;
        var viewServer = element.view_server;
        var optionalFilenameList = [];
        optionalFilenameList.push("template");
        var clientSeqNum = 0;

        var pullRequest = new PullRequest(id, viewServer, optionalFilenameList, clientSeqNum);
        var jsonPullRequest = JSON.stringify(pullRequest);

        var url = viewServer + '/viewserver/pull';

        $.ajax({
            url: url,
            type: 'POST',
            cache: false,
            data: jsonPullRequest,
            headers: { 'X-Username': username, 'X-Password': password },
            success: function (response, textStatus, jqXHR) {
                var parsedResponse = JSON.parse(response);
                var statusCode = parsedResponse.statusCode;
                var seqNum = parsedResponse.seqNum;

                var category = element.category;
                var metadata = JSON.parse(element.metadata);
                var metadaRowNum = metadata.rows;
                var metadaColumnNum = metadata.cols;

                if (statusCode == 200) {
                    var value = [];
                    //var optionalFileList = parsedResponse.optionalFileList;
                    var fileList = parsedResponse.fileList;
                    var rowNum, columnNum;
                    rowNum = columnNum = -1;

                    /*for (var i = 0; i < optionalFileList.length; i++) {
                        var file = optionalFileList[i];

                        var type = file.type;

                        var data = file.data;
                        var decodedData = window.atob(data);
                        var parsedDecodedData = JSON.parse(decodedData);

                        var excelData;
                        if (parsedDecodedData.Encrypt == false) {
                            excelData = window.atob(parsedDecodedData.data);

                        } else {

                        }
                    }*/

                    for (var i = 0; i < fileList.length; i++) {
                        var file = fileList[i];

                        var type = file.type;

                        var data = file.data;
                        var decodedData = window.atob(data);
                        var parsedDecodedData = JSON.parse(decodedData);

                        var excelData;
                        if (parsedDecodedData.Encrypt == false) {
                            excelData = window.atob(parsedDecodedData.data);

                        } else {

                        }

                        if (type == "FULL") {
                            var rows = excelData.split("\u001e");

                            if (rowNum == -1) {
                                rowNum = rows.length;
                            }
                            for (var j = 0; j < rows.length; j++) {
                                var row = rows[j];
                                var cells = row.split("\u001f");

                                value.push(cells);
                                if (columnNum == -1) {
                                    columnNum = cells.length;
                                }

                                /*if (category == "STOCK") {
                                    if (columnNum < metadaColumnNum) {
                                        for (var k = 0; k < (metadaColumnNum - columnNum) ; k++) {
                                            value[j][k + columnNum] = "";
                                        }
                                    }
                                }*/
                                /*if (columnNum == -1) {
                                    columnNum = cells.length;
                                }*/
                            }

                            /*if (category == "STOCK") {
                                rowNum = metadaRowNum;
                                columnNum = metadaColumnNum;
                            }*/

                        }
                        //TODO per SHARE
                        /*if (type == "DELTA") {
                            var delta = excelData.split("\u001c");

                            var initialRow = delta[0] - 1;
                            var initialColumn = delta[1] - 1;

                            var rows = delta[2].split("\u001e");

                            for (var j = 0; j < rows.length; j++) {
                                var row = rows[j];
                                var cells = row.split("\u001f");

                                for (var k = 0; k < cells.length; k++) {
                                    var cell = cells[k];

                                    value[initialRow + j][initialColumn + k] = cell;
                                }
                            }
                        }*/
                    }

                    Excel.run(function (ctx1) {
                        //a partire dall'address salvato in precedenza, ricavo la cella di partenza in cui copiare il risultato
                        var index = address.indexOf("!") + 1;
                        var wb = address.substring(0, index - 1);
                        var cell = address.substring(index);

                        //utilizzando rowIndex, columnIndex e le dimensioni del dato ottenuto ricavo l'address dell'ultima cella in cui andro' ad incollare i dati
                        var sheet = ctx1.workbook.worksheets.getItem(wb);
                        var firstCellRange = sheet.getRange(cell + ":" + cell);
                        var firstCell = sheet.getCell(rowIndex, columnIndex);
                        var lastCell = sheet.getCell(rowIndex + rowNum - 1, columnIndex + columnNum - 1);
                        lastCell.load('address');

                        /*var lastRow, lastColumn;
                        if (category == "STOCK") {
                            lastRow = sheet.getCell(rowIndex + rowNum - 1, columnIndex);
                            lastColumn = sheet.getCell(rowIndex, columnIndex + columnNum - 1);

                            lastRow.load('address');
                            lastColumn.load('address');
                        }*/

                        return ctx1.sync().then(function () {
                            //creo il range con i dati calcolati prima ed incollo il risultato ottenuto dal server
                            var range = sheet.getRange(cell + ":" + lastCell.address);
                            range.values = value;
                            var metadata = JSON.parse(element.metadata);
                            var isTable = metadata.is_table;
                            var has_headers = metadata.has_headers;
                            if (isTable) {
                                var table = ctx1.workbook.tables.add(address + ":" + lastCell.address, has_headers)
                            }

                            wb = address.substring(0, index - 1);
                            var lastCellRange = lastCell.address.substring(index);
                            var rangeAddress = address + ":" + lastCellRange;

                            addItemOnTable(rangeAddress, parsedResponse.seqNum, parsedResponse.fileList, parsedResponse.optionalFileList, -1);
                            $("#operationProgress").hide();
                            //se e' la vista della borsa inserisco i colori
                            /*if (category == "STOCK") {
                                sheet.getRange(cell + ":" + lastColumn.address).format.fill.color = "#D9D9D9";
                                sheet.getRange(cell + ":" + lastRow.address).format.fill.color = "#B4C6E7";
                                sheet.getRange(cell + ":" + cell).format.fill.color = "#92D050";
                            }*/

                        }).catch(function (error) {
                            app.showNotification('Error on pullView (1)');
                            $("#operationProgress").hide();
                        });
                    }).catch(function (error) {
                        app.showNotification('Error on pullView (2)');
                        $("#operationProgress").hide();
                    });

                    //se e' la vista della borsa aggiungo i bordi
                    /*if (category == "STOCK") {
                        Excel.run(function (ctx2) {
                            var index = address.indexOf("!") + 1;
                            var wb = address.substring(0, index - 1);
                            var cell = address.substring(index);

                            var sheet = ctx2.workbook.worksheets.getItem(wb);
                            var firstCellRange = sheet.getRange(cell + ":" + cell);
                            var firstCell = sheet.getCell(rowIndex, columnIndex);
                            var lastCell = sheet.getCell(rowIndex + rowNum - 1, columnIndex + columnNum - 1);
                            lastCell.load('address');

                            return ctx2.sync().then(function () {
                                var borderRange = sheet.getRange(cell + ":" + lastCell.address);
                                borderRange.format.borders.getItem('InsideHorizontal').style = 'Continuous';
                                borderRange.format.borders.getItem('InsideVertical').style = 'Continuous';
                                borderRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
                                borderRange.format.borders.getItem('EdgeLeft').style = 'Continuous';
                                borderRange.format.borders.getItem('EdgeRight').style = 'Continuous';
                                borderRange.format.borders.getItem('EdgeTop').style = 'Continuous';
                            }).catch(function (error) {
                                console.log('Error:', error);
                            });

                        }).catch(function (error) {
                            console.log("Error: " + error);
                            if (error instanceof OfficeExtension.Error) {
                                console.log("Debug info: " + JSON.stringify(error.debugInfo));
                            }
                        });
                    }*/

                } else {
                    app.showNotification('Error on pullView (3)');
                    $("#operationProgress").hide();
                }


            },
            error: function (jqXHR, textStatus, errorThrown) {
                app.showNotification('Error on pullView (4)');
                $("#operationProgress").hide();
            }
        });
    }

    function refreshView(event, seqNum) {
        var item = event.data.item;
        sssServer = $("#server").val();
        username = $("#username").val();
        password = $("#password").val();

        $("#operationProgress").show();

        var id = item.uuid;
        var viewServer = item.viewServer;
        var optionalFilenameList = [];
        optionalFilenameList.push("template");
        var clientSeqNum;
        if(seqNum != undefined) {
            clientSeqNum = seqNum;
        } else {
            clientSeqNum = item.seqNum;
        }
        

        var pullRequest = new PullRequest(id, viewServer, optionalFilenameList, clientSeqNum);
        var jsonPullRequest = JSON.stringify(pullRequest);

        var category = item.category;
        var metadaRowNum = item.rows;
        var metadaColumnNum = item.cols;

        var url = viewServer + '/viewserver/pull';

        $.ajax({
            url: url,
            type: 'POST',
            cache: false,
            data: jsonPullRequest,
            headers: { 'X-Username': username, 'X-Password': password },
            success: function (response, textStatus, jqXHR) {
                var parsedResponse = JSON.parse(response);
                var statusCode = parsedResponse.statusCode;
                var seqNum = parsedResponse.seqNum;

                if(clientSeqNum < seqNum) {
                    if (statusCode == 200) {
                        var value = [];
                        //var optionalFileList = parsedResponse.optionalFileList;
                        var fileList = parsedResponse.fileList;
                        var rowNum, columnNum;
                        rowNum = columnNum = -1;

                        /*for (var i = 0; i < optionalFileList.length; i++) {
                            var file = optionalFileList[i];
    
                            var type = file.type;
    
                            var data = file.data;
                            var decodedData = window.atob(data);
                            var parsedDecodedData = JSON.parse(decodedData);
    
                            var excelData;
                            if (parsedDecodedData.Encrypt == false) {
                                excelData = window.atob(parsedDecodedData.data);
    
                            } else {
    
                            }
                        }*/

                        for (var i = 0; i < fileList.length; i++) {
                            var file = fileList[i];

                            var type = file.type;

                            var data = file.data;
                            var decodedData = window.atob(data);
                            var parsedDecodedData = JSON.parse(decodedData);

                            var excelData;
                            if (parsedDecodedData.Encrypt == false) {
                                excelData = window.atob(parsedDecodedData.data);

                            } else {

                            }

                            if (type == "FULL") {
                                var rows = excelData.split("\u001e");

                                if (rowNum == -1) {
                                    rowNum = rows.length;
                                }
                                for (var j = 0; j < rows.length; j++) {
                                    var row = rows[j];
                                    var cells = row.split("\u001f");

                                    value.push(cells);
                                    if (columnNum == -1) {
                                        columnNum = cells.length;
                                    }
                                    /*if (category == "STOCK") {
                                        if (columnNum < metadaColumnNum) {
                                            for (var k = 0; k < (metadaColumnNum - columnNum) ; k++) {
                                                value[j][k + columnNum] = "";
                                            }
                                        }
                                    }*/
                                }

                                /*if (category == "STOCK") {
                                    rowNum = metadaRowNum;
                                    columnNum = metadaColumnNum;
                                }*/

                            }
                            //TODO per SHARE
                            /*if (type == "DELTA") {
                                var delta = excelData.split("\u001c");
    
                                var initialRow = delta[0] - 1;
                                var initialColumn = delta[1] - 1;
    
                                var rows = delta[2].split("\u001e");
    
                                for (var j = 0; j < rows.length; j++) {
                                    var row = rows[j];
                                    var cells = row.split("\u001f");
    
                                    for (var k = 0; k < cells.length; k++) {
                                        var cell = cells[k];
    
                                        value[initialRow + j][initialColumn + k] = cell;
                                    }
                                }
                            }*/
                        }

                        Excel.run(function (ctx) {
                            var index = item.rangeAddress.indexOf("!") + 1;
                            var sheetName = item.rangeAddress.substring(0, index - 1);
                            var rangeAddress = item.rangeAddress.substring(index);

                            var range;
                            /*if (listViewOnExcel[event.data.index].is_table) {
                                //range = ctx.workbook.tables.add(address + ":" + lastCell.address, has_headers)
                                range = ctx.workbook.tables.add(rangeAddress, listViewOnExcel[event.data.index].has_headers)
                            } else {
                                range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
                            }*/
                            range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
                            //var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
                            range.load('values');

                            return ctx.sync().then(function () {
                                range.values = value;
                                listViewOnExcel[event.data.index].seqNum = seqNum;

                                app.showNotification("Refresh of view completed");
                                $("#operationProgress").hide();
                            }).catch(function (error) {
                                app.showNotification('Error on refreshView (1)');
                                $("#operationProgress").hide();
                            });
                        }).catch(function (error) {
                            app.showNotification('Error on refreshView (2)');
                            $("#operationProgress").hide();
                        });

                    } else {
                        app.showNotification('Error on refreshView (3)');
                        $("#operationProgress").hide();
                    }
                } else {
                    if (clientSeqNum == seqNum) {
                        refreshView(event, seqNum - 1);
                    }
                    $("#operationProgress").hide();
                }

            },
            error: function (jqXHR, textStatus, errorThrown) {
                app.showNotification('Error on refreshView (4)');
                $("#operationProgress").hide();
            }
        });
    }
 
    function removeView(event) {
        var index = event.data.index;

        var viewElement = listViewOnExcel.splice(index, 1);
        checkBindingToRemove(viewElement[0].local_uuid);

        tableView();
    }
})();