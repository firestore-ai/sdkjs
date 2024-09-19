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
}

CRubyPr.prototype.Set_FromObject = function ( Obj )
{
    if (undefined !== Obj.type && null !== Obj.type) 
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
    NewPr.Align = Obj.Align;
    NewPr.Hps = Obj.Hps;
    NewPr.HpsRaise = Obj.HpsRaise;
    NewPr.HpsBaseText = Obj.HpsBaseText;
    NewPr.Lid = Obj.Lid;
    return NewPr;
}

CRubyPr.prototype.Write_ToBinary = function (Writer)
{
    Writer.WriteUInt8(this.Align);
    Writer.WriteLong(this.Hps);
    Writer.WriteLong(this.HpsRaise);
    Writer.WriteLong(this.HpsBaseText);
    Writer.WriteLong(this.Lid);
}

CRubyPr.prototype.Read_FromBinary = function (Reader)
{
    this.Align = Reader.ReadUInt8();
    this.Hps = Reader.ReadLong();
    this.HpsRaise = Reader.ReadLong();
    this.HpsBaseText = Reader.ReadLong();
    this.Lid = Reader.ReadLong();    
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

    AscCommon.g_oTableId.Add( this, this.Id );
}

ParaRuby.prototype = Object.create(AscWord.CRunElementBase.prototype);
ParaRuby.prototype.constructor = ParaRuby;

ParaRuby.prototype.Type = para_Ruby;

ParaRuby.prototype.SetParent = function(oParent )
{
    if (!oParent)
        return;

    if (oParent instanceof ParaRun)
        this.Parent = oParent.GetParagraph();
    else if (oParent instanceof Paragraph)
        this.Parent = oParent;
};
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
ParaRuby.prototype.Get_Run = function()
{
    return this.Get_Run();
}
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
// 测量, 计算元素的高和宽
ParaRuby.prototype.Measure = function(textMeasurer, textPr, infoMathText, paraR)
{
    // 高度
    var shaper =  AscWord.ParagraphTextShaper;
    shaper.ShapeRun(this.RubyText);
    shaper.ShapeRun(this.RubyBase);
    

    var baseMinMax = new CParagraphMinMaxContentWidth();
    this.RubyBase.RecalcMeasure();
    this.RubyBase.RecalculateMinMaxContentWidth(baseMinMax);
    var rubyMinMax = new CParagraphMinMaxContentWidth();
    this.RubyText.RecalcMeasure();
    this.RubyText.RecalculateMinMaxContentWidth(rubyMinMax);
    

    // 宽度

    this.Width = Math.max(baseMinMax.nCurMaxWidth, rubyMinMax.nCurMaxWidth);
    this.Height = baseMinMax.nMaxHeight + rubyMinMax.nMaxHeight;
    this.Ascent = 0;
    this.Descent = 0;
};
ParaRuby.prototype.getHeight = function () {
    return this.Height;
};
ParaRuby.prototype.GetWidth = function () {
    return this.Width;
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
    return true;
}

ParaRuby.prototype.IsText = function () {
    return false;
}

ParaRuby.prototype.Draw_Elements = function (PDSE)
{
    this.RubyText.Draw_Elements(PDSE);
    this.RubyBase.Draw_Elements(PDSE);    
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
}
ParaRuby.prototype.Read_FromBinary2 = function (Reader)
{
    this.Id = Reader.GetString2();
    this.Pr.Read_FromBinary(Reader);
    g_oTableId.Add(this, this.Id);
}

window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].ParaRuby = ParaRuby;

window['AscWord'] = window['AscWord'] || {};
window['AscWord'].ParaRuby = ParaRuby;
