using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace D365Batch.DownloadFileFieldData
{
    class StudentInfo
    {
        // テーブルロジック名
        public const string EntityLogicalName = "pas_tbl_student";
        // 列：学生番号
        public const string AttributeStudentNo = "pas_student_no";
        // 列：添付ファイル
        public const string AttributeFile = "pas_file";
        // 列：状態
        public const string AttributeStateCode = "statecode";
    }
}
