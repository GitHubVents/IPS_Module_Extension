using static AddFeatureContextMenu.Operations_with_IPS;

namespace AddFeatureContextMenu
{
    public struct MyStruct
    {
        public int Flag;
        public string Name;
        public string Designition;
        public string Path;
        public string RefAsmName;
        public IPSObjectTypes IPsDocType;
        public IPSObjectTypes IPsProductType;
        

        public MyStruct(int flag, string name, string designition, string path, string refAsmName, IPSObjectTypes docType, IPSObjectTypes productType)
        {
            Flag = flag;
            Name = name;
            Designition = designition;
            Path = path;
            RefAsmName = refAsmName;
            IPsDocType = docType;
            IPsProductType = productType;
        }
    }
}