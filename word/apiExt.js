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
     * Set custom xml     * 
     */
    asc_docs_api.prototype.asc_SetCustomXmlExt = function (uri, id, xml) {
        var document = this.private_GetLogicDocument();
        const encoder = new TextEncoder();
        return document.Update_CustomXml(uri, id, encoder.encode(xml));
    }

    asc_docs_api.prototype.asc_GetCustomXmlExt = function (id) {
        var decode = function (customXml) {
            const decoder = new TextDecoder();
            return {
                Uri: customXml.Uri,
                ItemId: customXml.ItemId,
                Content: String.fromCharCode.apply(String, Array.from(customXml.Content))
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
                if (customXml.Uri != undefined && customXml.Uri.includes(uri)) {
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


    /**
     * Make range by path
     * @param {string|number} beg - The beginning path or index
     * @param {string|number} end - The ending path or index
     * @returns {Range} - The created range
     */
    asc_docs_api.prototype.asc_MakeRangeByPath = function (beg, end) {
        if (typeof beg === 'number' && typeof end === 'number') {
            return this.GetDocument().GetRange().GetRange(beg, end);
        }

        var parsePath = function (path) {
            var indexes = path.match(/([-]?\d+)/g).map(e => parseInt(e));
            var currNode = this.private_GetLogicDocument();
            var positions = indexes.map(e => {
                if (currNode.Content.length === undefined) {
                    currNode = currNode.Content;
                }

                if (currNode.Content === undefined) {
                    return undefined;
                }

                if (e < 0) {
                    e = currNode.Content.length + e;
                }

                var position = {
                    Class: currNode,
                    Position: e
                };

                if (currNode.Content === undefined) {
                    debugger;
                }
                currNode = currNode.Content[e];

                return position;
            });
            positions = positions.filter(e => e !== undefined);

            return positions;
        }.bind(this);

        var startPos = parsePath(beg);
        var endPos = parsePath(end);

        // extend to range
        var ExtentToRun = function (isFirst, posArray) {
            var lastNode = posArray[posArray.length - 1];
            while (lastNode.Class.GetType == undefined || lastNode.Class.GetType() !== 39) {
                if (lastNode.Class.Content === undefined) {
                    break;
                }
                var next = {};
                next.Class = lastNode.Class.Content[lastNode.Position];
                if (!next.Class.Content.length) {
                    next.Class = next.Class.Content;
                }
                if (isFirst) {
                    next.Position = 0;
                } else {
                    next.Position = next.Class.Content.length - 1;
                }
                posArray.push(next);
                lastNode = next;
            }
            return posArray;
        };

        startPos = ExtentToRun(true, startPos);
        endPos = ExtentToRun(false, endPos);

        return this.GetDocument().GetRange(startPos, endPos);
    }

    // Regex search in range for word
    // @param {Range} range - The range to search
    // @param {pattern} pattern - The pattern to search
    // @return {Array} - The array of Range objects
    function marker_log(str, ranges) {
        let styledString = '';
        let currentIndex = 0;
        const styles = [];

        ranges.forEach(([start, end], index) => {
            // 添加高亮前的部分
            if (start > currentIndex) {
                styledString += '%c' + str.substring(currentIndex, start);
                styles.push('');
            }
            // 添加高亮部分
            styledString += '%c' + str.substring(start, end);
            styles.push('border: 1px solid red; padding: 2px');
            currentIndex = end;
        });

        // 添加剩余的部分
        if (currentIndex < str.length) {
            styledString += '%c' + str.substring(currentIndex);
            styles.push('');
        }

        console.log(styledString, ...styles);
    }

    function CalcTextPos(text_all, text_plain) {
        text_plain = text_plain.replace(/[\r]/g, '');
        var text_pos = new Array(text_all.length);
        var j = 0;
        for (var i = 0, n = text_plain.length; i < n; i++) {
            while (text_all[j] !== text_plain[i]) {
                text_pos[j] = i;
                j++;
            }
            text_pos[j] = i;
            j++;
        }
        return text_pos;
    }

    asc_docs_api.prototype.asc_RegexSearch = function (range, pattern, options = { log: false }) {
        // 用正则表达式实现
        // 自定义位置
        var text = range.GetText({ Math: false });
        var text_plain = range.GetText({ Math: false, Numbering: false });
        var text_pos = CalcTextPos(text, text_plain);

        var match;
        var matchRanges = [];
        var ranges = []
        while ((match = pattern.exec(text)) !== null) {
            var begPos = text_pos[match.index];
            var endPos = text_pos[match.index + match[0].length];
            ranges.push([match.index, match.index + match.length]);
            matchRanges.push(range.GetRange(begPos, endPos));
        }
        if (ranges.length > 0 && options.log) {
            marker_log(text, ranges);
        }
        return matchRanges;
    }




    // 将选中范围导出为ooxml
    asc_docs_api.prototype.asc_GenSelectionAsXml = function (options) {
        // copy selection to bin_data
        let bin_data = {
            data: "",
            // 返回的数据中class属性里面有binary格式的dom信息，需要删除掉
            pushData: function (format, value) {
                if (format === AscCommon.c_oAscClipboardDataFormat.Internal) {
                    this.data = value;
                }
            }
        };
        this.asc_CheckCopy(bin_data, AscCommon.c_oAscClipboardDataFormat.Internal);

        if (bin_data.data == "" || bin_data.data === undefined || bin_data.data === null) {
            console.log("asc_GenSelectionAsXml: bin_data is empty");
            if (options.callback != undefined)
                options.callback(undefined);
            return;
        }

        var oLogicDocument = this.private_GetLogicDocument();

        var isNoBase64 = false;
        var oAdditionalData = {};
        oAdditionalData["c"] = 'save';
        oAdditionalData["id"] = this.documentId;
        oAdditionalData["userid"] = this.documentUserId;
        oAdditionalData["tokenSession"] = this.CoAuthoringApi.get_jwt();
        oAdditionalData["outputformat"] = options.fileType;
        oAdditionalData["title"] = AscCommon.changeFileExtention(this.documentTitle, AscCommon.getExtentionByFormat(options.fileType), Asc.c_nMaxDownloadTitleLen);
        oAdditionalData["isNoBase64"] = isNoBase64;

        var dataContainer = { data: null, part: null, index: 0, count: 0 };
        dataContainer.data = bin_data.data.slice(8);

        //var oBinaryFileWriter = new AscCommonWord.BinaryFileWriter(oLogicDocument, undefined, undefined, options.compatible);
        //dataContainer.data = oBinaryFileWriter.Write(oAdditionalData["nobase64"]);


        let locale = this.asc_getLocale() || undefined;
        if (typeof locale === "string") {
            locale = Asc.g_oLcidNameToIdMap[locale];
        }
        oAdditionalData["lcid"] = locale;
        //oAdditionalData["withoutPassword"] = true;
        //oAdditionalData["inline"] = 1;
        var actionType = AscCommon.DownloadType.Download;
        var downloadType = actionType;

        this._downloadAsUsingServer(
            Asc.c_oAscAsyncAction.DownloadAs,
            options,
            oAdditionalData,
            dataContainer,
            actionType
        );

        return undefined;
    }


    asc_docs_api.prototype["asc_GetContentControlBoundingRectExt"] = asc_docs_api.prototype.asc_GetContentControlBoundingRectExt;
    asc_docs_api.prototype["asc_SetCustomXmlExt"] = asc_docs_api.prototype.asc_SetCustomXmlExt;
    asc_docs_api.prototype["asc_GetCustomXmlExt"] = asc_docs_api.prototype.asc_GetCustomXmlExt;
    asc_docs_api.prototype["asc_SetBiyueCustomDataExt"] = asc_docs_api.prototype.asc_SetBiyueCustomDataExt;
    asc_docs_api.prototype["asc_GetBiyueCustomDataExt"] = asc_docs_api.prototype.asc_GetBiyueCustomDataExt;

    asc_docs_api.prototype["asc_MakeRangeByPath"] = asc_docs_api.prototype.asc_MakeRangeByPath;
    asc_docs_api.prototype["asc_RegexSearch"] = asc_docs_api.prototype.asc_RegexSearch;
    asc_docs_api.prototype["asc_GenSelectionAsXml"] = asc_docs_api.prototype.asc_GenSelectionAsXml;

}(window, null));
