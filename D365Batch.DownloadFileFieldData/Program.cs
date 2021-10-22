using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Configuration;
using System.IO;

namespace D365Batch.DownloadFileFieldData
{
    class Program
    {
        static void Main()
        {
            try
            {
                // START...

                // ルートフォルダパス
                string rootFolderPath = ConfigurationManager.ConnectionStrings["RootFolderPath"].ConnectionString;
                // フォルダが存在かをチェック
                if (!Directory.Exists(rootFolderPath))
                {
                    // 存在しない時作成する
                    Directory.CreateDirectory(rootFolderPath);
                }

                // 認証を行い、組織サービスを取得する
                string ConnectionString = ConfigurationManager.ConnectionStrings["MyCDSServer"].ConnectionString;
                CrmServiceClient svc = new CrmServiceClient(ConnectionString);

                // 学生台帳の検索を行う
                QueryExpression studentQuery = new QueryExpression(StudentInfo.EntityLogicalName)
                {
                    // 結果列
                    ColumnSet = new ColumnSet(StudentInfo.AttributeStudentNo, StudentInfo.AttributeFile)
                };
                // 条件１：状態がアクティブ
                FilterExpression filter1 = new FilterExpression();
                filter1.AddCondition(new ConditionExpression(StudentInfo.AttributeStateCode, ConditionOperator.Equal, 0));
                // 条件２：添付ファイルが空白ではない
                FilterExpression filter2 = new FilterExpression();
                filter2.AddCondition(new ConditionExpression(StudentInfo.AttributeFile, ConditionOperator.NotNull));
                // 二つ条件の関係は且つ（AND）
                studentQuery.Criteria = new FilterExpression(LogicalOperator.And);
                studentQuery.Criteria.AddFilter(filter1);
                studentQuery.Criteria.AddFilter(filter2);
                // 検索を実行
                EntityCollection studentTable = svc.RetrieveMultiple(studentQuery);

                // 対象学生台帳をループ：開始
                foreach (Entity studentCol in studentTable.Entities)
                {
                    // ファイルダウンロード要求のインスタンス
                    InitializeFileBlocksDownloadRequest initializeFile = new InitializeFileBlocksDownloadRequest
                    {
                        // ファイル列のロジック名
                        FileAttributeName = StudentInfo.AttributeFile,
                        // 参照先エンティティ名及び参照先レコードGUIDでターゲットを指定
                        Target = new EntityReference(StudentInfo.EntityLogicalName, studentCol.Id)
                    };
                    // 要求して応答を取得
                    InitializeFileBlocksDownloadResponse initializeFileResponse = (InitializeFileBlocksDownloadResponse)svc.Execute(initializeFile);
                    Console.WriteLine($"ファイル名: {initializeFileResponse.FileName}");
                    Console.WriteLine($"ファイルサイズ (bytes): {initializeFileResponse.FileSizeInBytes}");
                    // 継続トークン
                    string fileContinuationToken = initializeFileResponse.FileContinuationToken;
                    // ファイル名
                    string fileName = initializeFileResponse.FileName;
                    // ファイルサイズ
                    long fileSize = initializeFileResponse.FileSizeInBytes;
                    // 開始位置
                    long offsetFrom = 0;
                    // 最大長さ
                    long blockLimitLength = 4 * 1024 * 1024;
                    // データ格納用
                    byte[] fileBytes = new byte[fileSize];

                    // ループ
                    while (offsetFrom < fileSize)
                    {
                        // ファイルダウンロード要求
                        DownloadBlockRequest blockRequest = new DownloadBlockRequest()
                        {
                            Offset = offsetFrom,
                            BlockLength = blockLimitLength,
                            FileContinuationToken = fileContinuationToken
                        };
                        DownloadBlockResponse blockResponse = (DownloadBlockResponse)svc.Execute(blockRequest);
                        blockResponse.Data.CopyTo(fileBytes, offsetFrom);
                        offsetFrom += blockLimitLength;
                    }

                    // ファイルの絶対パス
                    string childFolderPath = rootFolderPath +
                        "\\" +
                        studentCol.GetAttributeValue<string>(StudentInfo.AttributeStudentNo);

                    // フォルダが存在かをチェック
                    if (!Directory.Exists(childFolderPath))
                    {
                        // 存在しない時作成
                        Directory.CreateDirectory(childFolderPath);
                    }
                    // ローカルに保存
                    File.WriteAllBytes(Path.Combine(childFolderPath, fileName), fileBytes);
                }
                // 対象学生台帳をループ：終了

                // END...
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
    }
}
