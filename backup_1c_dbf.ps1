##################################################################################################
# Скрипт резервного копирования файловой базы данных 1С
# Дата создания: 08.11.2015
# 05/06/2016 fix: проводить ротацию только, если число имеющихся архивов больше $BackupRotationIntervalDays
##################################################################################################

# Логика работы
# Предполагается, что скрипт работает под учетной записью - BACKUP_OPERATOR с админскими правами 
# на локальной машине и чтение/изменение на удаленной.
# Проверяем, что нет активных пользователей БД (в этом случае все равно остается вероятность, что кто-то 
# подключится позднее, поэтому проводим операцию поздно ночью для минимизации вероятности).
# В случае наличия подключенных пользователей не производим бэкап (БД может быть не консистентна), 
# но отправляем нотификацию.
# В случае отсутствия таких пользователей создаем архив, в случае успеха делаем ротацию старых архивов. 
# По завершении работы отправляем нотификацию с детальным логом операции во вложении.
# При успешной отправке нотификации чистим рабочую директорию.
#



##################################################################################################
# Блок параметров
##################################################################################################

#Путь до исходной директории с базой для архивации:
$BackupSrcPath = "E:\1C.DB\current"

#Путь до целевой директории с архивами 
#(в ней не должно быть никаких других файлов, иначе rotation их может удалить)
$ArchiveDstPath = "\\SERVER\Backups\Базы 1С"

#Полный путь до рабочей директории (там же будет сохранен файл журнала задания)  
$WorkDirPath = "E:\BackupLogs"

#Путь до исполняемого файла архиватора 7-Zip
$SevenZipExecutablePath = "C:\Program files\7-Zip\7z.exe"

#Количество дней хранения архива (rotation)
$BackupRotationIntervalDays = 7

#Параметры для отправки почтовых нотификаций
$SmtpServer = "smtp-server" 
$MailFrom = "Backup_Operator@domain.ru" 
$MailTo = "it@domain.ru"


##################################################################################################
# Основная обработка
##################################################################################################
$ScriptLogPath=$WorkDirPath+"\"+"backup_1C_db"+"_$(Get-Date -format "yyyy-MM-dd").log"

Start-Transcript -path $ScriptLogPath

echo "--------------------------------------------------------"
echo "BackupSrcPath: $BackupSrcPath"
echo "ArchiveDstPath: $ArchiveDstPath"
echo "WorkDirPath: $WorkDirPath"
echo "SevenZipExecutablePath: $SevenZipExecutablePath"
echo "BackupRotationIntervalDays: $BackupRotationIntervalDays"
echo "--------------------------------------------------------"

#Проверяем, что нет открытых файлов БД
$SmbRawOutput = Get-SMBOpenfile
$SmbFailed = (!$?)

$SmbArray = @($SmbRawOutput)
$Pattern = $BackupSrcPath+"*"
$ListOpenedFiles = @($SmbArray | Where-Object { ($_.Path -Like "$Pattern") -and (Test-Path $_.Path -PathType Leaf) })

if ($SmbFailed)
{
	echo $SmbRawOutput
	
	#Не удалось получить информацию об открытых файлах БД.
	#Требуется вмешательство администратора для исправления проблемы
	$Encoding = [System.Text.Encoding]::UTF8
	$MailPriority = "High"
	$Subject = "[WARNING] Backup-1C"
	$Body = @"
Резервное копирование не проводилось, т.к. не удалось определить наличие открытых пользователями файлов базы данных. 
Требуется вмешательство администратора для исправления проблемы. Детальный лог операции во вложении.
"@
	
	Send-MailMessage -SmtpServer $SmtpServer -from $MailFrom -to ($MailTo) -Subject "$Subject" -Body "$Body" -Priority $MailPriority -Attachments $ScriptLogPath -Encoding $Encoding

	#Очищаем рабочую директорию
	if ($?)
	{
		Stop-Transcript
		Remove-Item -Path ($WorkDirPath+"\*") -Recurse	
	}
	else
	{
		Stop-Transcript
	}
}
elseif ($ListOpenedFiles.Count -ne 0)
{
	echo "Список открытых файлов:"
	$ListOpenedFiles | Format-Table -Wrap -Autosize -Property Path,ClientUserName,ClientComputerName
	<#
	foreach ($obj in $ListOpenedFiles)
	{
		echo "ClientComputerName - $($obj.ClientComputerName)"
		echo "ClientUserName - $($obj.ClientUserName)"
		echo "Path - $($obj.Path)"
		echo "`n"
	}
	#>
		
	#Есть открытые файлы БД, делаем только нотификацию без бэкапа
	$Encoding = [System.Text.Encoding]::UTF8
	$MailPriority = "High"
	$Subject = "[WARNING] Backup-1C"
	$Body = "Резервное копирование не проводилось из-за наличия открытых пользователями файлов базы данных. Детальный лог операции во вложении."
	
	Send-MailMessage -SmtpServer $SmtpServer -from $MailFrom -to ($MailTo) -Subject "$Subject" -Body "$Body" -Priority $MailPriority -Attachments $ScriptLogPath -Encoding $Encoding

	#Очищаем рабочую директорию
	if ($?)
	{
		Stop-Transcript
		Remove-Item -Path ($WorkDirPath+"\*") -Recurse	
	}
	else
	{
		Stop-Transcript
	}
}
else
{
	#TODO: попробовать поставить лочку на общий ресурс (Block-SmbShareAccess)
	
	echo "`nBackup started at: $(Get-Date -format "HH:mm dd-MM-yyyy")`n"

	$TargetBackupName = $ArchiveDstPath+"\"+"backup_1C_dbf"+"_$(Get-Date -format "yyyy-MM-dd").zip"
	echo "TargetBackupName: $TargetBackupName`n"

	#Создаем массив параметров для 7-Zip
	$Arg1="a"
	$Arg2="-tzip"
	$Arg3="-w"+$WorkDirPath
	$Arg4="-mx=5"
	$Arg5="-mmt=on"
	$Arg6="-ssw"
	$Arg7="-scsUTF-8"
	$Arg8="$TargetBackupName"
	$Arg9=$BackupSrcPath

	#Архивируем
	$SevenZipOutput = & $SevenZipExecutablePath ($Arg1,$Arg2,$Arg3,$Arg4,$Arg5,$Arg6,$Arg7,$Arg8,$Arg9) | Out-String
	echo $SevenZipOutput

	#Rotation 
	#(только в случае успешного создания бэкапа и если общее количество больше $BackupRotationIntervalDays, иначе вообще все поудаляем)
	if ($LASTEXITCODE -eq 0)
	{ 
		$TotalActualBackupsCount=(Get-ChildItem -Recurse -Force -Path "$ArchiveDstPath" | Measure-Object).Count
		if ($TotalActualBackupsCount -gt $BackupRotationIntervalDays)
		{ 
			echo "`nRotating old backups..."
			Get-ChildItem -Recurse -Force -Path "$ArchiveDstPath" | Sort-Object -property LastWriteTime | Select-Object -First ($TotalActualBackupsCount - $BackupRotationIntervalDays) | foreach {  
				$_ | Remove-Item
				$path = "$ArchiveDstPath"+"\"+"$_"
				if (Test-Path $path)
				{
					echo "[Fail] deleting - $path"
				}
				else
				{
					echo "[Success] deleting - $path"
				}
			}
		}
	}

	echo "`nBackup finished at: $(Get-Date -format "HH:mm dd-MM-yyyy")`n"

	#Отправляем нотификацию
	$Encoding = [System.Text.Encoding]::UTF8
	$MailPriority = "Normal"
	$Subject = "[OK] Backup-1C"
	$Body = "Резервное копирование завершено успешно. Детальный лог операции во вложении."
	if ($LASTEXITCODE -ne 0)
	{
		$MailPriority = "High"
		$Subject = "[FAILED] Backup-1C"
		$Body = "Не удалось завершить резервное копирование. Детальный лог операции во вложении."
	}

	Send-MailMessage -SmtpServer $SmtpServer -from $MailFrom -to ($MailTo) -Subject "$Subject" -Body "$Body" -Priority $MailPriority -Attachments $ScriptLogPath -Encoding $Encoding

	#Очищаем рабочую директорию
	if ($?)
	{
		Stop-Transcript
		Remove-Item -Path ($WorkDirPath+"\*") -Recurse	
	}
	else
	{
		Stop-Transcript
	}
}
