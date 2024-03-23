"use strict";
(function (window, builder) {
    /**
     * Base class
     * @global
     * @class
     * @name Api
     */
    var asc_docs_api = window["Asc"]["asc_docs_api"] || window["Asc"]["spreadsheet_api"];
    var c_oAscRevisionsChangeType = Asc.c_oAscRevisionsChangeType;
    var c_oAscSectionBreakType = Asc.c_oAscSectionBreakType;
    var c_oAscSdtLockType = Asc.c_oAscSdtLockType;
    var c_oAscAlignH = Asc.c_oAscAlignH;
    var c_oAscAlignV = Asc.c_oAscAlignV;

    /**
     * Extents the Api class
     */
    asc_docs_api.prototype.asc_GetContentControlBoundingRectExt = function (id) {
        var document = this.private_getDocument();
        var oElement = document.GetContentControlById(id);

        var ccRects = []
        for (var nPageIndex = 0, PageCount = this.Pages.length; nPageIndex < PageCount; nPageIndex++) {
            var Page = this.Pages[nPageIndex];
            for (var SectionIndex = 0, SectionsCount = Page.Sections.length; SectionIndex < SectionsCount; ++SectionIndex) {
                var PageSection = Page.Sections[SectionIndex];

                for (var ColumnIndex = 0, ColumnsCount = PageSection.Columns.length; ColumnIndex < ColumnsCount; ++ColumnIndex) {
                    var Column = PageSection.Columns[ColumnIndex];
                    var ColumnStartPos = Column.Pos;
                    var ColumnEndPos = Column.EndPos;

                    for (var ContentPos = ColumnStartPos; ContentPos <= ColumnEndPos; ++ContentPos) {
                        var oElement = this.Content[ContentPos];
                        if (Page.IsFlowTable(oElement) || Page.IsFrame(oElement)) {
                            continue;
                        }

                        if (oElement.GetType() === type_BlockLevelSdt) {
                            var rects = oElement.GetBoundingRect2();
                            for (var a = 0, n = rects.length; a < n; a++) {
                                var rect = rects[a];
                                if (rect && rect.Page == nPageIndex) { // fix: 避免重复
                                    ccRects.push(rects[a]);
                                }
                            }
                            // this.dumpBlockLevelSdt(oElement.Content.Content);
                        }
                    }
                }
            }
        }

        return ccRects;
    };

    /**
     * update custom xml
     */
    function updateCustomXml(document, uri, id, xml) {
        var customXmls = document.CustomXmls;
        for (var i = 0, n = customXmls.length; i < n; i++) {
            var customXml = customXmls[i];
            if (customXml.ItemId === id) {
                var oldContent = customXml.Content;
                customXml.Content = xml;
                return oldContent
            }
        }

        // if uri is string, convert uri = [uir]
        if (typeof uri === "string") {
            uri = [uri];
        }

        customXmls.push({
            Uri: uri,
            ItemId: id,
            Content: xml
        });
        return undefined;
    }

    /**
     * Set custom xml     * 
     */
    asc_docs_api.prototype.asc_SetCustomXmlExt = function (uri, id, xml) {
        var document = this.private_GetLogicDocument();
        const encoder = new TextEncoder();
        return updateCustomXml(document, uri, id, encoder.encode(xml));
    }

    asc_docs_api.prototype.asc_GetCustomXmlExt = function (id) {
        var decode = function (customXml) {
            const decoder = new TextDecoder();
            return {
                Uri: customXml.Uri,
                ItemId: customXml.ItemId,
                Content: String.fromCharCode(...customXml.Content)
            }
        }

        var document = this.private_GetLogicDocument();
        if (typeof id === "string") {
            var customXmls = document.CustomXmls;
            for (var i = 0, n = customXmls.length; i < n; i++) {
                var customXml = customXmls[i];
                if (customXml.ItemId === id) {
                    return decode(customXml);
                }
            }
            return undefined;
        } else if (typeof id === "number") {
            return decode(document.CustomXmls[id]);
        }

        return undefined;
    }

    // store js object to CDATA in custom xml 
    var encodeObjToXml = function (obj) {
        var xml = "<root><![CDATA[";
        xml += JSON.stringify(obj);
        xml += "]]></root>";
        return xml;
    }

    // decode js object from CDATA in custom xml
    var decodeObjFromXml = function (xml) {
        var obj = {};
        var parser = new DOMParser();
        var xmlDoc = parser.parseFromString(xml, "text/xml");
        var root = xmlDoc.getElementsByTagName("root")[0];
        root.textContent = root.textContent.trim();
        obj = JSON.parse(root.textContent);
        return obj;
    }

    // 获取biyue定制数据
    // 如果uuid为空，则返回所有的定制数据
    // 如果uuid不为空，则返回指定的定制数据
    asc_docs_api.prototype.asc_GetBiyueCustomDataExt = function (uuid) {
        var uri = "http://nicedoc/schema/question/1.0";
        // return all xmls          
        if (uuid == undefined) {
            var document = this.private_GetLogicDocument();
            var customXmls = document.CustomXmls;
            var result = [];
            for (var i = 0, n = customXmls.length; i < n; i++) {
                var customXml = this.asc_GetCustomXmlExt(i);
                // if uri in uri list                
                if (customXml.Uri.includes(uri)) {
                    result.push({ "ItemId": customXml.ItemId, "Content": decodeObjFromXml(customXml.Content) });
                }
            }
            return result;
        } else {
            var customXml = this.asc_GetCustomXmlExt(uuid);
            if (customXml !== undefined && customXml.Uri.includes(uri)) {
                return decodeObjFromXml(customXml.Content);
            } else {
                return undefined;
            }
        }
    }

    // 设置biyue定制数据
    // 如果uuid为空，则创建新的定制数据
    // 如果uuid不为空，则更新指定的定制数据
    // 返回uuid
    asc_docs_api.prototype.asc_SetBiyueCustomDataExt = function (uuid, data) {
        if (uuid == undefined) {
            uuid = window.AscCommon.CreateGUID();
        }
        var uris = ["http://nicedoc/schema/question/1.0"];
        var xml = encodeObjToXml(data);
        this.asc_SetCustomXmlExt(uris, uuid, xml);
        return uuid;
    }

    function ProcessNode(element, handler) {
        if (element === undefined || element === null) {
            return;
        }

        const cname = element.GetClassType()
        //const id = element.Id;

        let getText = function (element) {
            if (element.GetClassType() === "run") {
                return element.GetText();
            }
            return undefined;
        }

        handler.onStartElement(cname, getText(element));
        if (element.GetElementsCount != undefined) {
            for (var i = 0, n = element.GetElementsCount(); i < n; i++) {
                ProcessNode(element.GetElement(i), handler);
            }
        }
        handler.onEndElement(cname);
    }

    // 遍历文档dom结构
    asc_docs_api.prototype.asc_StartSaxParse = function (element, handler) {
        ProcessNode(element, handler);
    }

    asc_docs_api.prototype.asc_RunSaxTest = function () {
        var handler = {
            onStartElement: function (cname, text) {
                console.log("onStartElement: " + cname + " " + (text || ""));
            },
            onEndElement: function (cname) {
                console.log("onEndElement: " + cname + " ");
            }
        }

        this.asc_StartSaxParse(this.GetDocument(), handler);
    }

    asc_docs_api.prototype.asc_MakeRangeByPath = function (beg, end) {
        if (typeof beg === 'number' && typeof end === 'number')
            return Api.GetDocument().GetRange().GetRange(beg, end);
        // make range by path

        function toClass(node) {
            if (node.GetClassType() === "run") {
                return node.Run;
            } else if (node.GetClassType() === "paragraph") {
                return node.Paragraph;
            } else if (node.GetClassType() === "table") {
                return node.Table;
            } else if (node.GetClassType() === "cell") {
                return node.Cell;
            } else if (node.GetClassType() === "row") {
                return node.Row;
            } else if (node.GetClassType() === "document") {
                return node.Document;
            }
            return undefined;
        }

        var begIdxs = beg.match(/([-]?\d+)/g).map(e => parseInt(e));
        var endIdxs = end.match(/([-]?\d+)/g).map(e => parseInt(e));
        var startPos = []
        var endPos = []

        var currNode = this.private_GetLogicDocument();
        startPos = begIdxs.map(e => {
            // is content array
            if (!currNode.Content.length) {
                currNode = currNode.Content;                
            }

            if (e < 0) {
                e = currNode.Content.length  + e;
            }            

            var r  = {
                Class: currNode,
                Position: e
            };            
            currNode = currNode.Content[e];
            return r;
        });

        var currNode = this.private_GetLogicDocument();
        endPos = endIdxs.map(e => {
            if (!currNode.Content.length) {
                currNode = currNode.Content;                
            }

            if (e < 0) {
                e = currNode.Content.length  + e;
            }            

            var r =   {
                Class: currNode,
                Position: e
            };   
            currNode = currNode.Content[e];
            return r;            
        });

        return  this.GetDocument().GetRange(startPos, endPos);
    }


    asc_docs_api.prototype["asc_GetContentControlBoundingRectExt"] = asc_docs_api.prototype.asc_GetContentControlBoundingRectExt;
    asc_docs_api.prototype["asc_SetCustomXmlExt"] = asc_docs_api.prototype.asc_SetCustomXmlExt;
    asc_docs_api.prototype["asc_GetCustomXmlExt"] = asc_docs_api.prototype.asc_GetCustomXmlExt;
    asc_docs_api.prototype["asc_SetBiyueCustomDataExt"] = asc_docs_api.prototype.asc_SetBiyueCustomDataExt;
    asc_docs_api.prototype["asc_GetBiyueCustomDataExt"] = asc_docs_api.prototype.asc_GetBiyueCustomDataExt;
    asc_docs_api.prototype["asc_StartSaxParse"] = asc_docs_api.prototype.asc_StartSaxParse;

    asc_docs_api.prototype["asc_MakeRangeByPath"] = asc_docs_api.prototype.asc_MakeRangeByPath;


}(window, null));
