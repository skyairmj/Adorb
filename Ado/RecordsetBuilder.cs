using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ADODB;
using SD.LLBLGen.Pro.ORMSupportClasses;

namespace Utilities.Ado
{
    public class RecordsetBuilder<T> where T : EntityBase2, new()
    {
        private IEnumerable<T> entities;
        private T entity;
        private Field[] fields;
        private Action<RecordsetClass> headBuilder;
        private Action<RecordsetClass> headAdditionBuilder;
        private Action<RecordsetClass> bodyBuilder;
        private Action<RecordsetClass, T> rowBuilder;
        private Action<RecordsetClass, T> rowAdditionBuilder;

        private IEnumerable<T> Entities
        {
            get
            {
                if (entities == null && entity == null)
                {
                    entities = new T[] {};
                }
                return entities;
            }
        }

        public RecordsetBuilder<T> From(IEnumerable<T> entityParams)
        {
            entities = entityParams;
            return this;
        }

        public RecordsetBuilder<T> From(T entityParam)
        {
            if (entityParam == null) return this;
            entity = entityParam;
            entities = new[] {entity};
            return this;
        }

        public RecordsetBuilder<T> Select(params Field[] fieldParams)
        {
            fields = fieldParams;
            return this;
        }

        private void Initialise(RecordsetClass recordset)
        {
            if (recordset.RecordCount <= 0) return;
            recordset.Update(Missing.Value, Missing.Value);
            recordset.MoveFirst();
        }

        public RecordsetClass Build()
        {
            var recordset = new RecordsetClass();

            (headBuilder ?? BuildHead)(recordset);

            (headAdditionBuilder ?? (x => { }))(recordset);

            (bodyBuilder ?? BuildBody)(recordset);

            Initialise(recordset);

            return recordset;
        }

        public RecordsetBuilder<T> BuildHeadWith(Action<RecordsetClass> buildHeadAction)
        {
            headBuilder = buildHeadAction;
            return this;
        }

        public RecordsetBuilder<T> BuildHeadAdditionWith(Action<RecordsetClass> buildHeadAditionAction)
        {
            headAdditionBuilder = buildHeadAditionAction;
            return this;
        }

        public RecordsetBuilder<T> BuildBodyWith(Action<RecordsetClass> buildBodyAction)
        {
            bodyBuilder = buildBodyAction;
            return this;
        }

        public RecordsetBuilder<T> BuildRowWith(Action<RecordsetClass, T> buildRowAction)
        {
            rowBuilder = buildRowAction;
            return this;
        }

        public RecordsetBuilder<T> BuildRowAdditionWith(Action<RecordsetClass, T> buildRowAdditionAction)
        {
            rowAdditionBuilder = buildRowAdditionAction;
            return this;
        }

        private void BuildBody(RecordsetClass recordset)
        {
            recordset.Open(Missing.Value, Missing.Value, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockOptimistic,
                           1);
            foreach (var entity1 in Entities)
            {
                recordset.AddNew(Missing.Value, Missing.Value);
                (rowBuilder ?? BuildRow)(recordset, entity1);
                (rowAdditionBuilder ?? ((x, y) => { }))(recordset, entity1);
            }
        }

        private void BuildRow(_Recordset recordset, T _entity)
        {
            if (fields == null || fields.Count() == 0)
            {
                foreach (IEntityField2 field in _entity.Fields)
                {
                    BuildCell(recordset, _entity, field.Name, field.Name);
                }
            }
            else
            {
                foreach (var field in fields)
                {
                    BuildCell(recordset, _entity, field.Alias, field.Original);
                }
            }
        }

        private void BuildCell(_Recordset recordset, T _entity, string alias, string name)
        {
            var propertyInfo = typeof (T).GetProperty(name);
            var propertyValue = propertyInfo.GetValue(_entity, null);
            if (propertyInfo.PropertyType.Equals(typeof (DateTime))
                && DateTime.MinValue.Equals(propertyValue))
            {
                recordset.Fields[alias].Value = DBNull.Value;
            }
            else
            {
                recordset.Fields[alias].Value = propertyValue;
            }
        }

        private void BuildHead(RecordsetClass recordset)
        {
            T first;
            if (Entities != null && Entities.Count() > 0)
            {
                first = entities.First();
            }
            else
            {
                first = new T();
            }

            if (fields == null || fields.Count() == 0)
            {
                foreach (IEntityField2 field in first.Fields)
                {
                    var name = field.Name;
                    BuildHeadColumn(field, recordset, name);
                }
            }
            else
            {
                foreach (var field in fields)
                {
                    var name = field.Alias;
                    BuildHeadColumn(first.Fields[field.Original], recordset, name);
                }
            }
        }

        private void BuildHeadColumn(IEntityField2 field, RecordsetClass recordset, string name)
        {
            var dataTypeEnum = DataTypeEnum.adEmpty;
            var dataType = field.DataType;
            if (dataType.Equals(typeof (Int32)))
            {
                dataTypeEnum = DataTypeEnum.adInteger;
            }
            else if (dataType.Equals(typeof (string)))
            {
                dataTypeEnum = DataTypeEnum.adVarWChar;
            }
            else if (dataType.Equals(typeof (DateTime)))
            {
                dataTypeEnum = DataTypeEnum.adDBTimeStamp;
            }
            else if(dataType.Equals(typeof (Decimal) ))
            {
                dataTypeEnum = DataTypeEnum.adDouble;
            }
            else if (dataType.Equals(typeof(Double)))
            {
                dataTypeEnum = DataTypeEnum.adDouble;
            }
            else if (dataType.Equals(typeof(Boolean)))
            {
                dataTypeEnum = DataTypeEnum.adBoolean;
            }
            else
            {
                Console.WriteLine(dataType);
                Console.WriteLine(field.Name);
            }
            recordset.Fields.Append(name, dataTypeEnum, field.MaxLength, FieldAttributeEnum.adFldIsNullable,
                                    Missing.Value);
        }
    }
}
