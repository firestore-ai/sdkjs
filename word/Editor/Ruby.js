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

const para_Ruby = 0x0000000A;

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
function ParaRuby() {
    CParagraphContentWithParagraphLikeContent.call(this);

    this.Id                 = AscCommon.g_oIdCounter.GetNewId();
    this.Type               = para_Ruby;

    this.Pr                 = new CRubyPr();

    this.X                  = 0;
    this.Y                  = 0;

    this.FirstPage          = -1;
    
    this.bSelectionUse      = false;
    this.Paragraph          = null;

    this.NearPosArray       = [];   

    this.Width              = 0;
    this.WidthVisible       = 0;
    this.Height             = 0;
    this.Ascent             = 0;
    this.Descent            = 0;

    this.RubyText           = null;
    this.RubyBase           = null;

    AscCommon.g_oIdCounter.AddObject( this, this.Id );
}

ParaRuby.prototype.draw = function (x, y, pGraphics, PDSE)
{
    if (this.Paragraph) {
        this.Paragraph.draw(context, scale, page);
    }
}

ParaRuby.prototype.Draw_Elements = function (PDSE)
{
    if (this.bOneLine)
    {
        var X = PDSE.X;

        // Make_ShdColor

        for(var i=0; i < this.nRow; i++)
        {
            for(var j = 0; j < this.nCol; j++)
            {
                if(this.elements[i][j].IsJustDraw())
                {
                    var ctrPrp = this.Get_TxtPrControlLetter();

                    var Font = 
                    {
                        FontSize:   ctrPrp.FontSize,
                        FontFamily: {Name: ctrPrp.FontFamily.Name, Index: ctrPrp.FontFamily.Index},
                        Italic:     false,
                        Bold:       false
                    };

                    PDSE.Graphics.SetFont(Font);
                }
                this.elements[i][j].Draw_Elements(PDSE);
            }

        }
        PDSE.X = X + this.width;
    }
    else
    {
        var CurLine = PDSE.Line - this.StartLine;
        var CurRange = (0 === CurLine ? PDSE.Range - this.StartRange : PDSE.Range);

        var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange); 
        var EndPos = this.protected_GetRangeEndPos(CurLine, CurRange);

        for (var CurPos = StartPos; CurPos <= EndPos; CurPos++)
        {
            this.Content[CurPos].Draw_Elements(PDSE);
        }
    }
};



window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].ParaRuby = ParaRuby;
window['AscWord'].ParaRuby = ParaRuby;
