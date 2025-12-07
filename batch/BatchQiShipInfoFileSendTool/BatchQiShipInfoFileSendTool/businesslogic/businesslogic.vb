Imports System.Text.RegularExpressions
Imports Common
Imports Common.Log.Manager
Imports Common.Utilities
Imports System.Configuration

Public Class BusinessLogic
    Public Async Function UploadFile() As Task(Of Boolean)
        Dim fileLest As List(Of (DataPath As String, FinPath As String)) =
            Common.Utilities.FileTransferUtilities.GetReadyPairsHeadFin(
            folder:=""
            )

        If fileLest.Count = 0 Then
            ' 対象ファイルなし
            ConfigManager.Logger?.WriteLog(LogManager.LogLevel.INFO, "", "送信フォルダ配下にQI出荷情報ファイルが見つかりませんでした。")
            Return True
        End If

        For Each filePair In fileLest
            ' ファイルチェック処理
            If CheckFile(filePair) = False Then
                Continue For
            End If

            ' ファイルアップロード処理
            Dim uploadResult As Boolean = Await UploadFile(filePair.DataPath)
            If uploadResult = False Then
                ' アップロード失敗
                ConfigManager.Logger?.WriteLog(LogManager.LogLevel.ERROR, "", $"【アップロード失敗】{filePair.DataPath}")
                Return False
            End If
            ConfigManager.Logger?.WriteLog(LogManager.LogLevel.INFO, "", $"【アップロード成功】{filePair.DataPath}")

            ' アップロードファイル待避処理
            movefile(filePair)
        Next
        Return True
    End Function

    ''' <summary>
    ''' ファイルチェック処理
    ''' </summary>
    ''' <param name="filePair"></param>
    Private Function CheckFile(filePair As (DataPath As String, FinPath As String)) As Boolean

        'ファイル名チェック
        If IsValidQiShipFileName(filePair.DataPath) = False Then
            ConfigManager.Logger?.WriteLog(LogManager.LogLevel.INFO, "", "【フォーマット相違】処理対象外ファイルのためスキップ")
            Return False
        End If

        'アップロード済ファイルチェック
        If IsUpload(filePair.DataPath) = False Then
            ConfigManager.Logger?.WriteLog(LogManager.LogLevel.INFO, "", $"【アップロード済ファイル】{filePair.DataPath}")
            movefile(filePair)
            Return False
        End If

        Return True
    End Function

    Public Function IsValidQiShipFileName(filePath As String) As Boolean
        Dim fileName As String = FileUtilities.GetFileNameFromPath(filePath)

        ' ファイル名形式：
        ' qi_ship_info_SMCxxxxx_YYYYMMDDHHMMSS.csv
        Dim pattern As String = "^qi_ship_info_(SMC.{8})_(\d{14})\.csv$"
        Dim match As Match = Regex.Match(fileName, pattern, RegexOptions.IgnoreCase)

        If Not match.Success Then
            ConfigManager.Logger?.WriteLog(LogManager.LogLevel.INFO, "", $"【ファイル名不正】{fileName}")
            Return False
        End If

        ' --- ① キー番号（出荷番号） ---
        If IsShKaNo(match.Groups(1).Value) = False Then
            Return False
        End If

        ' --- ② 日付（YYYYMMDDHHMMSS）が妥当か ---
        Dim dtStr As String = match.Groups(2).Value
        Dim dt As DateTime
        If Not DateTime.TryParseExact(dtStr, "yyyyMMddHHmmss",
                                  Nothing,
                                  Globalization.DateTimeStyles.None,
                                  dt) Then
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' 待避フォルダにファイルが存在するかチェックする。
    ''' </summary>
    ''' <param name="filePath"></param>
    ''' <returns></returns>
    Private Function IsUpload(filePath As String) As Boolean
        Dim fileName As String = FileUtilities.GetFileNameFromPath(filePath)
        Dim bkupFoler As String = ConfigurationManager.AppSettings("BKUP_FOLDER")
        Dim checkFilePath As String = Common.Utilities.FileUtilities.CombinePath(bkupFoler, fileName)

        Return FileUtilities.IsFileExist(checkFilePath)
    End Function


    Private Async Function UploadFile(filePath As String) As Task(Of Boolean)

        ' S3接続オブジェクト作成
        Dim accessKey As String = ConfigurationManager.AppSettings("S3_ACCESS_KEY")
        Dim secretKey As String = ConfigurationManager.AppSettings("S3_SECRET_KEY")
        Dim region As String = ConfigurationManager.AppSettings("S3_REGION")
        Dim bucketName As String = ConfigurationManager.AppSettings("S3_BUCKET_NAME")

        Dim s3 As New Common.Aws.S3(
                accessKey:=accessKey,
                secretKey:=secretKey,
                region:=region,
                bucketName:=bucketName)

        ' ファイルアップロード
        Dim key As String = String.Format(ConfigurationManager.AppSettings("S3_KEY"), FileUtilities.GetFileNameFromPath(filePath))
        Dim result As Boolean = Await s3.UploadS3File(key, filePath)

        Return result
    End Function

    ''' <summary>
    ''' 待避フォルダにファイルを移動する。
    ''' </summary>
    ''' <param name="filePair"></param>
    Private Sub MoveFile(filePair As (DataPath As String, FinPath As String))
        MoveFile(filePair.DataPath)
        MoveFile(filePair.FinPath)
    End Sub

    ''' <summary>
    ''' 待避フォルダにファイルを移動する
    ''' </summary>
    ''' <param name="filePath"></param>
    Private Sub MoveFile(filePath As String)
        ' 処理済みファイル移動処理
        Dim FileName As String = FileUtilities.GetFileNameFromPath(filePath)
        Dim bkupFolder As String = ConfigurationManager.AppSettings("BKUP_FOLDER")
        Dim bkupFilePath As String = FileUtilities.CombinePath(bkupFolder, FileName)

        Common.Utilities.FileUtilities.MoveFile(
                srcPath:=filePath,
                destPath:=bkupFilePath)
        ConfigManager.Logger?.WriteLog(LogManager.LogLevel.INFO, "", $"【待避フォルダ移動（移動元）】{filePath}")
        ConfigManager.Logger?.WriteLog(LogManager.LogLevel.INFO, "", $"【待避フォルダ移動（移動先）】{bkupFilePath}")
    End Sub
End Class
