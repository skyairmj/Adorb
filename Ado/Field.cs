using SD.LLBLGen.Pro.ORMSupportClasses;

namespace Utilities.Ado
{
    public class Field
    {
        private string alias;

        public Field(string original)
        {
            Original = original;
        }

        public Field(EntityField2 entityField) : this(entityField.Name)
        {
        }

        public EntityField2 EntityField { get; private set; }
        public string Original { get; private set; }

        public string Alias
        {
            get { return alias ?? Original; }
            private set { alias = value; }
        }

        public static implicit operator Field(string field)
        {
            return new Field(field);
        }

        public static implicit operator Field(EntityField2 field)
        {
            return new Field(field);
        }

        public Field As(string alias)
        {
            Alias = alias;
            return this;
        }
    }
}
