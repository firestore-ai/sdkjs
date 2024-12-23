/*
 * (c) Copyright Ascensio System SIA 2010-2023
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";
(function (window, builder) {
    var asc_docs_api = window["Asc"]["asc_docs_api"] || window["Asc"]["spreadsheet_api"];
    var c_oAscRevisionsChangeType = Asc.c_oAscRevisionsChangeType;
    var c_oAscSectionBreakType = Asc.c_oAscSectionBreakType;
    var c_oAscSdtLockType = Asc.c_oAscSdtLockType;
    var c_oAscAlignH = Asc.c_oAscAlignH;
    var c_oAscAlignV = Asc.c_oAscAlignV;

    // 获取段落的边界    
    asc_docs_api.prototype.asc_GetParagraphBoundingRect = function(sId, page) {
        var oLogicDocument = this.private_GetLogicDocument();
        if (!oLogicDocument)
            return null;

        const oParagraph = AscCommon.g_oTableId.Get_ById(sId);
        if (!oParagraph || !oParagraph.GetContentBounds)
            return null;

        if (page === undefined) 
        {
            page = oParagraph.GetStartPageAbsolute()
        }

        var oBounds = oParagraph.GetContentBounds(page);
        if (!oBounds)
            return null;

        var oRect = {
            X: oBounds.Left,
            Y: oBounds.Top,
            W: oBounds.Bottom - oBounds.Top,
            H: oBounds.Right - oBounds.Left,
            Transform : oParagraph.Get_ParentTextTransform()
        }

        var nX, nY, nW, nH;

        var oTransform = oRect.Transform;
        if (oTransform)
        {
            var nX0 = oTransform.TransformPointX(oRect.X, oRect.Y);
            var nY0 = oTransform.TransformPointY(oRect.X, oRect.Y);
            var nX1 = oTransform.TransformPointX(oRect.X + oRect.W, oRect.Y);
            var nY1 = oTransform.TransformPointY(oRect.X + oRect.W, oRect.Y);
            var nX2 = oTransform.TransformPointX(oRect.X + oRect.W, oRect.Y + oRect.H);
            var nY2 = oTransform.TransformPointY(oRect.X + oRect.W, oRect.Y + oRect.H);
            var nX3 = oTransform.TransformPointX(oRect.X, oRect.Y + oRect.H);
            var nY3 = oTransform.TransformPointY(oRect.X, oRect.Y + oRect.H);

            nX = Math.min(nX0, nX1, nX2, nX3);
            nY = Math.min(nY0, nY1, nY2, nY3);
            nW = Math.max(nX0, nX1, nX2, nX3) - nX;
            nH = Math.max(nY0, nY1, nY2, nY3) - nY;
        }
        else
        {
            nX = oRect.X;
            nY = oRect.Y;
            nW = oRect.W;
            nH = oRect.H;
        }

        return {
            Page: page,
            X0: nX,
            Y0: nY,
            X1: nX + nW,
            Y1: nY + nH
        };
    }

    let MeasureNumberingTextWidth = function(oPara, nCharCount) {                
        var oNumbering = oPara.Parent.GetNumbering();
        var oTheme = oPara.GetTheme();
        var oNumTextPr = oPara.GetNumberingTextPr();
        var sId = oPara.Numbering.Internal.FinalNumId;
        var nLvl =oPara.Numbering.Internal.FinalNumLvl;

        var oNum =  oNumbering.GetNum(sId);
        var oLvl    = oNum.GetLvl(nLvl);
        var arrText = oLvl.GetLvlText();
        var dKoef   = oNumTextPr.VertAlign !== AscCommon.vertalign_Baseline ? AscCommon.vaKSize : 1;

	    let g_oTextMeasurer =  AscCommon.g_oTextMeasurer;
        g_oTextMeasurer.SetTextPr(oNumTextPr, oTheme);
        g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII, dKoef);
        
        let NumPr = oPara.GetNumPr();
        var arrNumInfo = oPara.Parent.CalculateNumberingValues(oPara, NumPr, true) || [];
        var oNumInfo = arrNumInfo[0] || [];

        var Width = 0;
    
        for (var nTextIndex = 0, nTextLen = arrText.length; nTextIndex < nTextLen && nTextIndex < nCharCount; ++nTextIndex)
        {
            switch (arrText[nTextIndex].Type)
            {
                //case numbering_lvltext_Text:
                case 1:
                {
                    let strValue  = arrText[nTextIndex].Value;
                    let codePoint = strValue.charCodeAt(0);
                    let curCoef   = dKoef;
    
                    let info;
                    if ((info = oNum.ApplyTextPrToCodePoint(codePoint, oNumTextPr)))
                    {
                        curCoef *= info.FontCoef;
                        codePoint = info.CodePoint;
                        strValue  = String.fromCodePoint(codePoint);
                    }
    
                    var FontSlot = AscWord.GetFontSlotByTextPr(codePoint, oNumTextPr);
    
                    g_oTextMeasurer.SetFontSlot(FontSlot, curCoef);
    
                    Width += g_oTextMeasurer.Measure(strValue).Width;
    
                    break;
                }
                //case numbering_lvltext_Num:
                case 2:
                {
                    g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII, dKoef);
                    var langForTextNumbering = oNumTextPr.Lang;
    
                    var nCurLvl = arrText[nTextIndex].Value;
                    var T = "";
    
                    if (nCurLvl < oNumInfo.length)
                        T = oNum.private_GetNumberedLvlText(nCurLvl, oNumInfo[nCurLvl], oLvl.IsLegalStyle() && nCurLvl < nLvl, langForTextNumbering);
    
                    for (var iter = T.getUnicodeIterator(); iter.check(); iter.next())
                    {
                        var CharCode = iter.value();
                        Width += g_oTextMeasurer.MeasureCode(CharCode).Width;
                    }
    
                    break;
                }
            }
        }
        return Width;
    }

    // 获取段落编号的边界
    // @param {string} sId - 段落ID
    // @param {number} charCount - 字符元素数，如果不传则返回整个编号的边界，如果传了则返回指定字符数的边界，如果%1也只算一个元素
    asc_docs_api.prototype.asc_GetParagraphNumberingBoundingRect = function(sId, charCount) {
        var oLogicDocument = this.private_GetLogicDocument();
        if (!oLogicDocument)
            return null;

        const oParagraph = AscCommon.g_oTableId.Get_ById(sId);
        if (!oParagraph || !oParagraph.GetContentBounds)
            return null;

        if (oParagraph.GetNumPr() === undefined) 
            return null;
        
        var oNum = oParagraph.Numbering;
        if (!oNum)
            return null;

        var oBounds = oParagraph.GetMulLineBounds(oNum.Page, 1);
        if (!oBounds)
            return null;                        
        
        var oRect = {
            Page: oNum.Page,
            X: oBounds.Left,
            Y: oBounds.Top,
            W: oNum.Width,
            H: oBounds.Bottom - oBounds.Top,
            Transform : oParagraph.Get_ParentTextTransform()
        }

        if (charCount !== undefined && charCount > 0) 
        {
           oRect.W = MeasureNumberingTextWidth(oParagraph, charCount);
        }            
       
        return {
                Page: oRect.Page,
                X0: oRect.X,
                Y0: oRect.Y,
                X1: oRect.X + oRect.W,
                Y1: oRect.Y + oRect.H
        };        
    }

    /**
     * Extents the Api class
     */
    asc_docs_api.prototype.asc_GetContentControlBoundingRectExt = function (sId) {
        var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return null;

		var oContentControl = oLogicDocument.GetContentControl(sId);
		if (!oContentControl)
			return null;

		var aRect = oContentControl.GetBoundingRect2();

		if (!aRect || aRect.length <= 0)
			return null;

        var rects = aRect.map((oRect) => {
		    var nX, nY, nW, nH;
		    var oTransform = oRect.Transform;
		    if (oTransform)
		    {
                var nX0 = oTransform.TransformPointX(oRect.X, oRect.Y);
                var nY0 = oTransform.TransformPointY(oRect.X, oRect.Y);
                var nX1 = oTransform.TransformPointX(oRect.X + oRect.W, oRect.Y);
                var nY1 = oTransform.TransformPointY(oRect.X + oRect.W, oRect.Y);
                var nX2 = oTransform.TransformPointX(oRect.X + oRect.W, oRect.Y + oRect.H);
                var nY2 = oTransform.TransformPointY(oRect.X + oRect.W, oRect.Y + oRect.H);
                var nX3 = oTransform.TransformPointX(oRect.X, oRect.Y + oRect.H);
                var nY3 = oTransform.TransformPointY(oRect.X, oRect.Y + oRect.H);

                nX = Math.min(nX0, nX1, nX2, nX3);
                nY = Math.min(nY0, nY1, nY2, nY3);
                nW = Math.max(nX0, nX1, nX2, nX3) - nX;
                nH = Math.max(nY0, nY1, nY2, nY3) - nY;
            }
		    else
		    {
                nX = oRect.X;
                nY = oRect.Y;
                nW = oRect.W;
                nH = oRect.H;
		    }

            return {
                Page: oRect.Page,
                X0: nX,
                Y0: nY,
                X1: nX + nW,
                Y1: nY + nH
            };
        });

        return rects;
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
                Content: decoder.decode(Uint8Array.from(customXml.Content))
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

    // 计算文档网格可以设置的设置的最大行数
    asc_docs_api.prototype.asc_GetMaxGridRows = function (height) {
        var nPageHeight = AscCommon.MMToTwip(height);
        var nGridHeight = 285;
        var nGridRows = Math.floor(nPageHeight / nGridHeight);
        return nGridRows;
    }

    // 根据行数计算一行高度
    asc_docs_api.prototype.asc_CalcLinePitch = function (height, row) {
        if (row == 0) {
            return height;
        }
        var nPageHeight = AscCommon.MMToTwip(height);
        return Math.floor(nPageHeight / row)        
    }

    asc_docs_api.prototype.asc_SetLines = function(lines)
    {
        var oLogicDocument = this.private_GetLogicDocument();
        if (!oLogicDocument)
            return;

        var CurPos = oLogicDocument.CurPos.ContentPos;
        var SectPr = oLogicDocument.SectionsInfo.Get_SectPr(CurPos).SectPr;

        var DocGrid = SectPr.DocGrid;
    
        const height = SectPr.GetContentFrameHeight();
        const maxLines = asc_GetMaxGridRows(height);
        if (lines > maxLines) {
            lines = maxLines;
        }

        var linePitch = asc_CalcLinePitch(height, lines);
        
        SectPr.SetDocGridLinePitch(linePitch);
    }

    asc_docs_api.prototype.asc_ReplaceWithRuby = function(range, ruby)
    {
        range.Select();
        range.Delete();
        asc_InsertRuby(ruby);
    }

    asc_docs_api.prototype.asc_InsertRuby = function(ruby)
    {

    }

    asc_docs_api.prototype.asc_SetTab = function(name)
	{
		this.sendEvent("asc_onSetTab", name);
	};

    asc_docs_api.prototype.asc_OpenPlugin = function(guid)
	{
		this.sendEvent("asc_onOpenPlugin", guid);
	};
	

    asc_docs_api.prototype["asc_GetParagraphBoundingRect"] = asc_docs_api.prototype.asc_GetParagraphBoundingRect;
    asc_docs_api.prototype["asc_GetParagraphNumberingBoundingRect"] = asc_docs_api.prototype.asc_GetParagraphNumberingBoundingRect;

    asc_docs_api.prototype["asc_GetContentControlBoundingRectExt"] = asc_docs_api.prototype.asc_GetContentControlBoundingRectExt;
    asc_docs_api.prototype["asc_SetCustomXmlExt"] = asc_docs_api.prototype.asc_SetCustomXmlExt;
    asc_docs_api.prototype["asc_GetCustomXmlExt"] = asc_docs_api.prototype.asc_GetCustomXmlExt;
    asc_docs_api.prototype["asc_SetBiyueCustomDataExt"] = asc_docs_api.prototype.asc_SetBiyueCustomDataExt;
    asc_docs_api.prototype["asc_GetBiyueCustomDataExt"] = asc_docs_api.prototype.asc_GetBiyueCustomDataExt;

    asc_docs_api.prototype["asc_MakeRangeByPath"] = asc_docs_api.prototype.asc_MakeRangeByPath;
    asc_docs_api.prototype["asc_RegexSearch"] = asc_docs_api.prototype.asc_RegexSearch;
    asc_docs_api.prototype["asc_GenSelectionAsXml"] = asc_docs_api.prototype.asc_GenSelectionAsXml;

    asc_docs_api.prototype["asc_SetTab"] = asc_docs_api.prototype.asc_SetTab;
    asc_docs_api.prototype["asc_OpenPlugin"] = asc_docs_api.prototype.asc_OpenPlugin;
    

}(window, null));
