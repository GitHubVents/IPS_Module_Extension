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

        // 0 - part, 1 - asm, 2 - drw
        public MyStruct(int flag, string name, string designition, string path, string refAsmName, int docType)
        {
            Flag = flag;
            Name = name;
            Designition = designition;
            Path = path;
            DefineTypes(docType, out IPsDocType, out IPsProductType);
            RefAsmName = refAsmName;
        }

        private static void DefineTypes(int docType, out IPSObjectTypes IPsDocType, out IPSObjectTypes IPsProductType)
        {

            switch (docType)
            {
                case 0:
                    IPsDocType = IPSObjectTypes.DocSldPart;
                    IPsProductType = IPSObjectTypes.ProductPart;
                    break;

                case 1:
                    IPsDocType = IPSObjectTypes.DocSldAssembly;
                    IPsProductType = IPSObjectTypes.ProductASM;
                    break;

                case 2:
                    IPsDocType = IPSObjectTypes.DocSldDrw;
                    IPsProductType = IPSObjectTypes.Empty;//???????????????????????????????????????????????????
                    break;

                default:
                    IPsDocType = IPSObjectTypes.Empty;
                    IPsProductType = IPSObjectTypes.Empty;
                    break;

            }

        }
    }

}
