"use strict";
(function(window, builder)
{
	/**
	 * Base class
	 * @global
	 * @class
	 * @name Api
	 */
	var Api = window["Asc"]["asc_docs_api"] || window["Asc"]["spreadsheet_api"];
	var c_oAscRevisionsChangeType = Asc.c_oAscRevisionsChangeType;
	var c_oAscSectionBreakType    = Asc.c_oAscSectionBreakType;
	var c_oAscSdtLockType         = Asc.c_oAscSdtLockType;
	var c_oAscAlignH         = Asc.c_oAscAlignH;
	var c_oAscAlignV         = Asc.c_oAscAlignV;

    /**
     * Extents the Api class
     */
    Api.prototype.GetContentControlBoundingRectExt = function (id)
    {
        var document = this.private_getDocument();
        var oElement = this.GetContentControlById(id);

        var ccRects = []
        for (var nPageIndex = 0, PageCount = this.Pages.length; nPageIndex < PageCount; nPageIndex++) {
            var Page = this.Pages[nPageIndex];
            for (var SectionIndex = 0, SectionsCount = Page.Sections.length; SectionIndex < SectionsCount; ++SectionIndex) {
                var PageSection = Page.Sections[SectionIndex];

                for (var ColumnIndex = 0, ColumnsCount = PageSection.Columns.length; ColumnIndex < ColumnsCount; ++ColumnIndex) {
                    var Column         = PageSection.Columns[ColumnIndex];
                    var ColumnStartPos = Column.Pos;
                    var ColumnEndPos   = Column.EndPos;

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





}(window, null));
