f = function(){ 
    var a = "160;BgAAADYANgAyAB4AAwAAAAAA+uAAAgUAAAAAAAAAAAAAAAAAAAMAAACghgEAAQAB////AQAAAf///wEAAAEAAAAItIgOAAQAAAAABgAAADYAMwAzAAACAAH64AACBQAAAAAAAADb0QoAAAAAAwAAAKCGAQABAAH///8BAAAB////AQAAAQAAAAi0iA4ABAAAAAAGAAAANgAzADMAAAIAAQ=="
    var memoryData = AscCommon.Base64.decode(a, true, undefined, undefined)
    var reader= new AscCommon.FT_Stream2(memoryData, memoryData.length);
    var ClassName = reader.GetString2();
    var Class = AscCommon.g_oTableId.Get_ById("660");
    var id = reader.GetLong();
    var fChangesClass= AscDFH.changesFactory[id];

    var oChange = new fChangesClass(Class)
    debugger;
    oChange.ReadFromBinary(reader)
}