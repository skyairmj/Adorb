using SD.LLBLGen.Pro.ORMSupportClasses;

namespace Utilities.Ado
{
    public static class FieldExtension
    {
        public static Field As(this string str, string alias)
        {
            return new Field(str).As(alias);
        }

        public static Field As(this EntityField2 entityField2, string alias)
        {
            return new Field(entityField2).As(alias);
        }
    }
}
