using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.DocumentValidator;

namespace Counsel_System2.ValidationRule
{
    public class CounselRowValidatorFactory : IRowValidatorFactory
    {
        #region IRowValidatorFactory 成員

        IRowVaildator IRowValidatorFactory.CreateRowValidator(string typeName, System.Xml.XmlElement validatorDescription)
        {
            switch (typeName.ToUpper())
            {
                case "STUDDATACHECKQUIZDATASTUDENTNUMBERRVAL":
                    return new RowValidator.StudDataCheckQuizDataStudentNumberRVal();
                case "CHECKTEACHERNAMERVAL":
                    return new RowValidator.CheckTeacherNameRVal();
                case "COUNSELSTUDCHECKSTUDENTNUMBERSTATUSVAL":
                    return new RowValidator.StudCheckStudentNumberStatusVal();
                case "STUDCHECKCLASSSEATNOSTATUSVAL":
                    return new RowValidator.StudCheckClassSeatNoStatusVal();
                default:
                    return null;
            }
        }

        #endregion
    }
}
