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

// RubyText 是一种文本格式，用于在文本中添加注音或解释。
// 它通常用于日语和汉语等语言中，以帮助读者理解文本的含义或者标记读音。
// RubyText 由两部分组成：Ruby 文本和基文本。
// Ruby 文本是注音或解释，基文本是实际的文本内容。
// RubyText 通常以红色显示，以区别于普通文本。

// Ruby 对齐方式的枚举值
var c_oAscRubyInterfaceSettingsAlign = {
    Center              : 0, // 居中对齐
    DistributeLetter    : 1, // 按字母分布
    DistributeSpace     : 2, // 按空格分布
    Left                : 3, // 左对齐
    Right               : 4, // 右对齐
    RightVertical       : 5  // 垂直右对齐
};

// Import
var g_oTextMeasurer = AscCommon.g_oTextMeasurer;
var History = AscCommon.History;
var c_oAscRevisionsChangeType = Asc.c_oAscRevisionsChangeType;

/**
 * 创建一个新的 Ruby 属性对象
 * @constructor
 */
function CRubyPr() {
    // Ruby 文本的对齐方式，默认为居中对齐
    this.Align  = align_Center;
    
    // Ruby 文本的字号大小（半点）
    this.Hps = 0;
    
    // Ruby 文本相对于基准文本的上升高度（半点）
    this.HpsRaise = 0;
    
    // 基准文本的字号大小（半点）
    this.HpsBaseText = 0;
    
    // 语言标识符
    this.Lid = 0;

    this.Type               = para_RubyPr;
}
ParaRuby.prototype.Type = para_RubyPr;
CRubyPr.prototype.Set_FromObject = function ( Obj )
{
    if (undefined !== Obj.Type && null !== Obj.Type) 
    {
        this.Align = Obj.Align;
        this.Hps = Obj.Hps;
        this.HpsRaise = Obj.HpsRaise;
        this.HpsBaseText = Obj.HpsBaseText;
        this.Lid = Obj.Lid;
    }
}

CRubyPr.prototype.Copy = function (Obj) 
{
    var NewPr = new CRubyPr();
    NewPr.Align = this.Align;
    NewPr.Hps = this.Hps;
    NewPr.HpsRaise = this.HpsRaise;
    NewPr.HpsBaseText = this.HpsBaseText;
    NewPr.Lid = this.Lid;
    return NewPr;
}

CRubyPr.prototype.Write_ToBinary = function (Writer)
{    
    Writer.WriteByte(this.Align);
    Writer.WriteLong(this.Hps);
    Writer.WriteLong(this.HpsRaise);
    Writer.WriteLong(this.HpsBaseText);
    Writer.WriteLong(this.Lid);
}

CRubyPr.prototype.Read_FromBinary = function (Reader)
{
    this.Align = Reader.GetByte();
    this.Hps = Reader.GetLong();
    this.HpsRaise = Reader.GetLong();
    this.HpsBaseText = Reader.GetLong();
    this.Lid = Reader.GetLong();    
}


/**
 * 
 * @constructor
 * @extends {CParagraphContentWithParagraphLikeContent}
 */
function ParaRuby(documentContent, Parent) {
    AscWord.CRunElementBase.call(this);

    this.Id                 = AscCommon.g_oIdCounter.Get_NewId();
    this.Type               = para_Ruby;

    this.Pr                 = new CRubyPr();

    this.X                  = 0;
    this.Y                  = 0;

    this.FirstPage          = -1;
    
    this.bSelectionUse      = false;
    this.Paragraph          = Parent;
    this.Run                = null;
    this.DocumentContent    = documentContent;
    this.Parent             = Parent;

    this.NearPosArray       = [];   

    this.Width              = 0;
    this.WidthVisible       = 0;
    this.Height             = 0;
    this.Ascent             = 0;
    this.Descent            = 0;

    this.RubyText           = null;
    this.RubyBase           = null;    
};
ParaRuby.prototype = Object.create(AscWord.CRunElementBase.prototype);
ParaRuby.prototype.constructor = ParaRuby;
ParaRuby.prototype.Type = para_Ruby;
ParaRuby.prototype.AddToTable = function()
{
    AscCommon.g_oTableId.Add( this, this.Id );
};
ParaRuby.prototype.Copy = function(oPr)
{
    var newRuby = new ParaRuby(this.documentContent, this.Paragraph);

    newRuby.Pr = this.Pr.Copy(oPr);

    newRuby.RubyText = this.RubyText.Copy(oPr);
    newRuby.RubyBase = this.RubyBase.Copy(oPr);    

    newRuby.AddToTable();

    return newRuby;
};
ParaRuby.prototype.SetParent = function(oParent )
{
    if (!oParent)
        return;

    if (oParent instanceof ParaRun)
    {
        this.Run = oParent;
        this.Parent = oParent.GetParagraph();        
    }
    else if (oParent instanceof Paragraph)
        this.Parent = oParent;
};
ParaRuby.prototype.SetParagraph = function(oParagarph)
{
    this.Paragraph = oParagarph;
    if (this.RubyText)
        this.RubyText.Paragraph = this.Paragraph;
    if (this.RubyBase)
        this.RubyBase.Paragraph = this.Paragraph;        
}
ParaRuby.prototype.GetParent = function()
{
    return this.Parent;
};
ParaRuby.prototype.Get_Type = function()
{
    return this.Type;
};
ParaRuby.prototype.Get_Paragraph = function()
{
    return this.Get_ParentParagraph();
};
ParaRuby.prototype.GetDocumentContent = function()
{
    const oParagraph = this.GetParagraph();
    let oDocumentContent = (oParagraph ? oParagraph.GetParent() : null);
    if (oDocumentContent && oDocumentContent.IsBlockLevelSdtContent())
    {
        oDocumentContent = oDocumentContent.Parent.Parent;
    }
    return oDocumentContent;
}
ParaRuby.prototype.Get_Run = function()
{
    var oParagraph = this.Get_ParentParagraph();
    // TODO: 这里需要修改
    if (oParagraph)
        return oParagraph.Get_DrawingObjectRun(this.Id);
    return null;
};
ParaDrawing.prototype.GetParagraph = function()
{
	return this.Get_ParentParagraph();
};
ParaRuby.prototype.IsInline = function() {
    return true;
}
ParaRuby.prototype.Recalculate_Range = function(PRS, ParaPr, Depth)
{
    this.Recalculate_Reset(this.Run.StartRange, this.Run.StartLine);
    var X = PRS.X;
    var Y = PRS.Y;
    var XEnd = PRS.XEnd;
    var YEnd = PRS.YEnd;

    this.RubyText.Recalculate_Range(PRS, ParaPr, Depth);
    var rubyX = PRS.X;
    PRS.X = X;
    this.RubyBase.Recalculate_Range(PRS, ParaPr, Depth);
    PRS.X = Math.max(rubyX, PRS.X);
};
ParaRuby.prototype.Recalculate_Reset = function(StartRange, StartLine)
{
    this.RubyText.Recalculate_Reset(StartRange, StartLine);
    this.RubyBase.Recalculate_Reset(StartRange, StartLine);

};
ParaRuby.prototype.GetAllFontNames = function(AllFonts)
{
    if (this.RubyText)
        this.RubyText.Get_AllFontNames(AllFonts);    
    if (this.RubyBase)
        this.RubyBase.Get_AllFontNames(AllFonts);
};
ParaRuby.prototype.CreateDocumentFontMap = function(Map) 
{
    if (this.RubyText)
        this.RubyText.Create_FontMap(Map);    
    if (this.RubyBase)
        this.RubyBase.Create_FontMap(Map);
};
ParaRuby.prototype.GetFontSlot = function(oTextPr)
{
    return this.RubyBase.GetFontSlotInRange(0, this.RubyBase.Content.length) |
        this.RubyText.GetFontSlotInRange(0, this.RubyText.Content.length);	
}
// 测量, 计算元素的高和宽
ParaRuby.prototype.Measure = function(textMeasurer, textPr, infoMathText, paraR)
{
    this.RubyText.Paragraph = this.Paragraph;
    this.RubyBase.Paragraph = this.Paragraph;        
    
    // 高度
    AscWord.ParagraphTextShaper.ShapeRun(this.RubyText);
    AscWord.ParagraphTextShaper.ShapeRun(this.RubyBase);

    var baseMinMax = new CParagraphMinMaxContentWidth();
    this.RubyBase.RecalcMeasure();
    this.RubyBase.RecalculateMinMaxContentWidth(baseMinMax);
    var rubyMinMax = new CParagraphMinMaxContentWidth();
    this.RubyText.RecalcMeasure();
    this.RubyText.RecalculateMinMaxContentWidth(rubyMinMax);
    

    // 宽度
    this.rubyWidth = rubyMinMax.nCurMaxWidth;
    this.baseWidth = baseMinMax.nCurMaxWidth;

    this.SetWidth(Math.max(baseMinMax.nCurMaxWidth, rubyMinMax.nCurMaxWidth));
    this.Height = baseMinMax.nMaxHeight + rubyMinMax.nMaxHeight;
    this.Ascent = 0;
    this.Descent = 0;
    this.SetWidthVisible(this.GetWidth());
};
ParaRuby.prototype.getHeight = function () {
    return this.Height;
};

ParaRuby.prototype.Draw = function (x, y, pGraphics, PDSE) {
    this.draw(x, y, pGraphics, PDSE);
    pGraphics.End_Command();
}

ParaRuby.prototype.draw = function (x, y, pGraphics, PDSE)
{    
    this.Draw_Elements(PDSE);
}
ParaRuby.prototype.IsDrawing = function () {
    return false;
}
ParaRuby.prototype.IsText = function () {
    return false;
}
ParaRuby.prototype.Draw_Elements = function (PDSE)
{
    // 保存当前的 X 和 Y 坐标
    var X = PDSE.X;
    var Y = PDSE.Y;

    // 获取 Ruby 属性
    var pr = this.Pr;
    // 计算 Ruby 文本的垂直偏移量（将 Twips 转换为毫米）
    var yOffset = AscCommon.TwipsToMM(pr.HpsRaise*10);

    // 调整 Y 坐标以绘制 Ruby 文本（向上偏移）
    PDSE.Y = Y - yOffset;

    var width = this.GetWidth();
    switch (this.Pr.Align)
    {
    case AscCommon.ruby_align_Center:
    default:
        PDSE.X += (width - this.rubyWidth)/2;
        break;
    case AscCommon.ruby_align_Left:
        break;
    case AscCommon.ruby_align_Right:
        PDSE.X += (width - this.rubyWidth);
        break;
    }
    
    // 绘制 Ruby 文本
    this.RubyText.Draw_Elements(PDSE);
    
    // 重置 X 坐标到原始位置
    PDSE.X = X;
    // 重置 Y 坐标到原始位置，准备绘制基础文本
    PDSE.Y = Y;    
    // 绘制基础文本
    this.RubyBase.Draw_Elements(PDSE);    

    // 最后，将 X 坐标重置到原始位置
    // 这确保了后续绘制操作从正确的位置开始
    PDSE.X = X;
};
ParaRuby.prototype.Write_ToBinary = function (Writer)
{
    Writer.WriteLong(this.Type);
    Writer.WriteString2(this.Id);
}
ParaRuby.prototype.Write_ToBinary2 = function (Writer)
{
    Writer.WriteLong(AscDFH.historyitem_type_Ruby);
    Writer.WriteString2(this.Id);    
    this.Pr.Write_ToBinary(Writer);
    this.RubyBase.Write_ToBinary2(Writer);    
    this.RubyText.Write_ToBinary2(Writer);
}
ParaRuby.prototype.Read_FromBinary2 = function (Reader)
{
    this.Id = Reader.GetString2();
    this.Pr.Read_FromBinary(Reader);
    Reader.GetLong(); // bypass type
    this.RubyBase = new ParaRun();
    this.RubyBase.Read_FromBinary2(Reader);
    Reader.GetLong(); // bypass type
    this.RubyText = new ParaRun();
    this.RubyText.Read_FromBinary2(Reader);    
    g_oTableId.Add(this, this.Id);
}

window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].ParaRuby = ParaRuby;

window['AscWord'] = window['AscWord'] || {};
window['AscWord'].ParaRuby = ParaRuby;
