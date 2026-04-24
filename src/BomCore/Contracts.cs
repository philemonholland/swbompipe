using System.IO;

namespace BomCore;

public interface IAssemblyReader
{
    AssemblyReadResult ReadActiveAssembly();
}

public interface ISelectedComponentPropertyReader
{
    IReadOnlyList<PropertyValue> ReadSelectedComponentProperties();
}

public interface IBomExporter
{
    void Export(BomResult result, Stream output);
}
