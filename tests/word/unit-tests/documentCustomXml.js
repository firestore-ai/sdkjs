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

$(function () {
    let logicDocument = AscTest.CreateLogicDocument();
	let styleManager  = null;


	QUnit.module("Unit-tests for CChangesDocumentCustomXml");

    QUnit.test("Test xml -> undefined:", (assert)=>
    {
        const encoder = new TextEncoder();
        var customXml = {
            ItemId: "123",
            Uri : [
                "http://schemas.openxmlformats.org/officeDocument/2006/customXml",
            ],
            Content : encoder.encode("<root><element>value</element></root>")
        }
        let change = new CChangesDocumentCustomXml(logicDocument, customXml, undefined);
        var memory = new AscCommon.CMemory(true)
        change.WriteToBinary(memory);
        
        let change2 = new CChangesDocumentCustomXml(logicDocument);
        var stream = new AscCommon.FT_Stream2(memory.data, memory.pos);
        change2.ReadFromBinary(stream);

        
        assert.deepEqual(change2.New, change.New, "New");
        assert.deepEqual(change2.Old, change.Old, "Old");        
        
    });


	QUnit.test("Test undefined -> xml:", (assert)=>
	{
        const encoder = new TextEncoder();
        var customXml = {
            ItemId: "123",
            Uri : [
                "http://schemas.openxmlformats.org/officeDocument/2006/customXml",
            ],
            Content : encoder.encode("<root><element>value</element></root>")
        }
		let change = new CChangesDocumentCustomXml(logicDocument, undefined, customXml);
        var memory = new AscCommon.CMemory(true)
        change.WriteToBinary(memory);
        
        let change2 = new CChangesDocumentCustomXml(logicDocument);
        var stream = new AscCommon.FT_Stream2(memory.data, memory.pos);
        change2.ReadFromBinary(stream);

        
        assert.deepEqual(change2.New, change.New, "New");
        assert.deepEqual(change2.Old, change.Old, "Old");        

	});


    QUnit.test("Test old -> new:", function (assert)
	{
        const encoder = new TextEncoder();
        var customXml = {
            ItemId: "123",
            Uri : [
                "http://schemas.openxmlformats.org/officeDocument/2006/customXml",
            ],
            Content : encoder.encode("<root><element>value</element></root>")
        }

        var newCustomXml = {
            ItemId: "123",
            Uri : [
                "http://schemas.openxmlformats.org/officeDocument/2006/customXml",
            ],
            Content : encoder.encode("<root><element>new</element></root>")
        }

		let change = new CChangesDocumentCustomXml(logicDocument, customXml, newCustomXml);
        var memory = new AscCommon.CMemory(true)
        change.WriteToBinary(memory);
        
        let change2 = new CChangesDocumentCustomXml(logicDocument);
        var stream = new AscCommon.FT_Stream2(memory.data, memory.pos);
        change2.ReadFromBinary(stream);

        
        assert.deepEqual(change2.New, change.New, "New");
        assert.deepEqual(change2.Old, change.Old, "Old");        
    });

	
});
