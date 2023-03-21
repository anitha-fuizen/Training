var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'News1WebPartStrings';
import styles from './components/News1.module.scss';
import { SPHttpClient } from '@microsoft/sp-http';
import "@pnp/sp/sputilities";
//import { IHttpClientOptions, HttpClientResponse, HttpClient } from '@microsoft/sp-http';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { getSP } from './components/pnpConfig';
var GetListItemFromSharePointListWebPart = /** @class */ (function (_super) {
    __extends(GetListItemFromSharePointListWebPart, _super);
    function GetListItemFromSharePointListWebPart(props) {
        return _super.call(this) || this;
        // this.send_Email=this.send_Email.bind(this);
    }
    GetListItemFromSharePointListWebPart.prototype._getListData = function () {
        return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('Training')/Items?$select=Title,Description", SPHttpClient.configurations.v1)
            .then(function (response) {
            return response.json();
            console.log(response.json());
        });
    };
    GetListItemFromSharePointListWebPart.prototype._renderListAsync = function () {
        var _this = this;
        if (Environment.type === EnvironmentType.SharePoint ||
            Environment.type === EnvironmentType.ClassicSharePoint) {
            this._getListData()
                .then(function (response) {
                _this._renderList(response.value);
                console.log(response.value);
            }).catch(function (err) { console.log(err); });
        }
    };
    GetListItemFromSharePointListWebPart.prototype._getData = function (useraddress) {
        var _this = this;
        return this.context.spHttpClient
            .get(this.context.pageContext.web.absoluteUrl +
            "/_api/web/lists/GetByTitle('Employee List')/Items?$select=Reportingmanager/EMail&$expand=Reportingmanager&$filter = employeename/EMail eq " + useraddress, SPHttpClient.configurations.v1).then(function (response) { return __awaiter(_this, void 0, void 0, function () {
            var resItems;
            return __generator(this, function (_a) {
                resItems = response.json();
                return [2 /*return*/, resItems];
            });
        }); });
    };
    GetListItemFromSharePointListWebPart.prototype.send_Email = function (title) {
        return __awaiter(this, void 0, void 0, function () {
            var _sp_1, addressString, emailString_1, mymanager, e_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        _sp_1 = getSP(this.context);
                        return [4 /*yield*/, _sp_1.utility.getCurrentUserEmailAddresses()];
                    case 1:
                        addressString = _a.sent();
                        //let mymanager:string= await _sp.web.lists.getByTitle("EmployeeDetails").items.getItemByStringId() 
                        //const items:any=   await _sp.web.lists.getByTitle("Employee List").items.select("Reportingmanager/Title").expand("Reportingmanager")();
                        console.log(addressString);
                        emailString_1 = null;
                        mymanager = this._getData(addressString);
                        mymanager.then(function (x) {
                            var obj = x.value;
                            obj.forEach(function (x) {
                                emailString_1 = x.Reportingmanager.EMail;
                                console.log(emailString_1);
                                _sp_1.utility.sendEmail({
                                    To: [emailString_1],
                                    Subject: "Request for" + title,
                                    Body: "Iam interested in " + title,
                                    AdditionalHeaders: {
                                        "content-type": "text/html"
                                    },
                                });
                            });
                            return emailString_1;
                        });
                        //return items
                        window.alert("success");
                        console.log("emailsend");
                        return [3 /*break*/, 3];
                    case 2:
                        e_1 = _a.sent();
                        console.log(e_1);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    GetListItemFromSharePointListWebPart.prototype._renderList = function (items) {
        return __awaiter(this, void 0, void 0, function () {
            var html, x, listContainer, buttons, labels, _loop_1, i;
            var _this = this;
            return __generator(this, function (_a) {
                html = '<table border=2 width=80% style="font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;>';
                x = 0;
                html += "<th>Title</th><th>Description</th><th>Apply</th>";
                console.log(items);
                items.forEach(function (item) {
                    x = x + 1;
                    html += "<tr>\n\n       <td><label> ".concat(item.Title, "</label></td>\n       <td> ").concat(item.Description, "</td>\n\n       <td> <button id=\"nominate_btn").concat(x, "\" class=\"nominate1\">Nominate</button></td><br><br>\n       \n        \n        </tr> ");
                });
                html += "</table>";
                listContainer = this.domElement.querySelector('#BindspListItems');
                listContainer.innerHTML = html;
                buttons = listContainer.getElementsByTagName("BUTTON");
                labels = listContainer.getElementsByTagName("LABEL");
                if (buttons) {
                    _loop_1 = function (i) {
                        buttons[i].addEventListener('click', function (e) { return _this.send_Email(labels[i].textContent); });
                    };
                    for (i = 0; i < buttons.length; i++) {
                        _loop_1(i);
                    }
                }
                return [2 /*return*/];
            });
        });
    };
    GetListItemFromSharePointListWebPart.prototype.render = function () {
        this.domElement.innerHTML = "\n      <div class={styles.sharepointframe}>\n    <div class={ styles.container }>\n      <div class={ styles.row }>\n        <div class={ column }>\n        <span class=\"".concat(styles.title, "\"></span>\n          \n          </div>\n          <br/>\n          <br/>\n          <br/>\n          <div id=\"BindspListItems\" />\n          </div>\n          </div>\n           \n          </div>");
        this._renderListAsync();
        this._getData("sindhu");
    };
    Object.defineProperty(GetListItemFromSharePointListWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    GetListItemFromSharePointListWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return GetListItemFromSharePointListWebPart;
}(BaseClientSideWebPart));
export default GetListItemFromSharePointListWebPart;
//# sourceMappingURL=News1WebPart.js.map