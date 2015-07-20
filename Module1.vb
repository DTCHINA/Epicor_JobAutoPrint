Imports Epicor.Mfg.Core
Imports Epicor.Mfg.BO
Imports Epicor.Mfg.Lib
Imports System.IO


Module Module1

    Sub Main()
        'Delcare our Business Objects
        Dim _session As Epicor.Mfg.Core.Session

        Dim _job As Epicor.Mfg.BO.JobEntry
        Dim _plant As Plant
        Dim _sessioMod As SessionMod
        'Connect to Epicor
        _session = New Epicor.Mfg.Core.Session("epicorapm", "bvgy789", "AppServerDC://Server8:9401", Epicor.Mfg.Core.Session.LicenseType.DataCollection)


        Console.WriteLine("---------------START------------" & Now() & "-------------")
        Dim coList As List(Of String) = New List(Of String)()
        coList.Add("CIC68322")
        coList.Add("10")
        coList.Add("20")
        Try
            For Each s As String In coList

                _session.CompanyID = s
                _sessioMod = New SessionMod(_session.ConnectionPool)
                '_sessioMod.SetCompany("10", "", "", "", "", "", "", "", "")
                _plant = New Plant(_session.ConnectionPool)
                Dim morePages As Boolean
                Dim planDs As PlantDataSet = _plant.GetRows("Plant.Company='" & s & "'", "", "", 0, 0, morePages)

                For Each plantRow As PlantDataSet.PlantRow In planDs.Plant
                    '***************************************************
                    ' Query Job Head for 4 days and set mass print flag 
                    ' Query Job Head for 8 Days in MX ONLY!
                    ' All jobs, all Plants
                    '***************************************************
                    _session.PlantID = plantRow.Plant
                    _job = New Epicor.Mfg.BO.JobEntry(_session.ConnectionPool)
                    Dim joDS As JobEntryDataSet
                    Dim moreRecords As Boolean
                    Dim jobHeadWhereClause As String
                    Dim DaysAhead As Double
                    DaysAhead = 6

                    'DO NOT CHANGE THIS LINE
                    If _session.CompanyID & _session.PlantID = "10MfgSys" Then
                        DaysAhead = 6
                    End If

                    'DO NOT CHANGE THIS LINE
                    If Weekday(Now) = 1 Then DaysAhead = DaysAhead + 1 ' Sunday Start
                    If Weekday(Now) = 3 Then DaysAhead = DaysAhead + 2 ' Tuesday Start
                    If Weekday(Now) = 4 Then DaysAhead = DaysAhead + 2 ' Wednesday Start
                    If Weekday(Now) = 5 Then DaysAhead = DaysAhead + 2 ' Thursday Start
                    If Weekday(Now) = 6 Then DaysAhead = DaysAhead + 2 ' Friday Start
                    If Weekday(Now) = 7 Then DaysAhead = DaysAhead + 2 ' Saturday Start

                    'DO NOT CHANGE THIS LINE
                    If _session.CompanyID & _session.PlantID = "CIC6832206" Then
                        DaysAhead = 8
                    End If

                    'DO NOT CHANGE THIS LINE
                    If _session.CompanyID & _session.PlantID = "CIC6832205" Then
                        DaysAhead = 9
                    End If

                    'DO NOT CHANGE THIS LINE
                    'CHANGE THESE LINES
                    'If _session.CompanyID & _session.PlantID = "CIC6832205" Then 'TX Plant'
                    'If _session.CompanyID & _session.PlantID = "CIC6832206" Then 'MX Plant"
                    If _session.CompanyID & _session.PlantID = "CIC68322MfgSys" Then 'CT Plant'
                        'If _session.CompanyID & _session.PlantID = "10Mfgsys" Then 'UK Plant'
                        DaysAhead = 9
                    End If



                    Console.WriteLine("Mass Print Flags --" & _session.CompanyID & " " & _session.PlantID & "--Days Ahead--" & DaysAhead)

                    '***************************************************
                    'Query Job Head 
                    '***************************************************
                    jobHeadWhereClause = String.Format("JobHead.Plant='{1}' AND JobHead.JobFirm = True AND JobHead.JobComplete = False AND JobHead.StartDate < '{0}' AND JobHead.TravelerLastPrinted = ? BY JobHead.JobNum", Date.Today.AddDays(DaysAhead), _session.PlantID)
                    joDS = _job.GetRows(jobHeadWhereClause, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 100, 0, moreRecords)
                    UpdateJobs1(joDS, _job)  'Update the Jobs

                    While moreRecords  'While the DB has more, do it all over again
                        jobHeadWhereClause = String.Format("JobHead.Plant='{2}' AND JobHead.JobFirm = True AND JobHead.JobComplete = False AND JobHead.StartDate < '{0}' AND JobHead.TravelerLastPrinted = ? AND JobHead.JobNum > '{1}' BY JobHead.JobNum", Date.Today.AddDays(DaysAhead), joDS.JobHead(joDS.JobHead.Count - 1).JobNum, _session.PlantID)
                        moreRecords = False
                        joDS = _job.GetRows(jobHeadWhereClause, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 100, 0, moreRecords)
                        UpdateJobs1(joDS, _job)
                    End While
                    '***************************************************

                Next

                '***************************************************'***************************************************'***************************************************'***************************************************

                For Each plantRow As PlantDataSet.PlantRow In planDs.Plant
                    '***************************************************
                    ' Query Job Head for 6 days 
                    ' Texas jobs with part class LP
                    '***************************************************
                    _session.PlantID = plantRow.Plant
                    _job = New Epicor.Mfg.BO.JobEntry(_session.ConnectionPool)
                    Dim joDS As JobEntryDataSet
                    Dim moreRecords As Boolean
                    Dim jobHeadWhereClause As String
                    Dim DaysAhead As Double
                    DaysAhead = 0 ' DO NOT CHANGE THIS NUMBER OR BAD THINGS WILL HAPPEN.  

                    If _session.CompanyID & _session.PlantID = "CIC6832206" Then
                        'As of 1/21/2014 Dwight Marshel, this is not changeing as of yet.
                        DaysAhead = 10 ' only change this number as need

                    End If

                    Console.WriteLine("LPs ----- Mass Print Flags --" & _session.CompanyID & " " & _session.PlantID & "--Days Ahead--" & DaysAhead)

                    '***************************************************
                    'Query Job Head 
                    '***************************************************
                    jobHeadWhereClause = String.Format("JobHead.Plant='{1}' AND JobHead.JobFirm = True AND JobHead.JobComplete = False AND JobHead.StartDate < '{0}' AND JobHead.TravelerLastPrinted = ? AND JobHead.Prodcode='{2}' BY JobHead.JobNum", Date.Today.AddDays(DaysAhead), _session.PlantID, "DC")
                    joDS = _job.GetRows(jobHeadWhereClause, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 100, 0, moreRecords)
                    UpdateJobs1(joDS, _job)  'Update the Jobs

                    While moreRecords  'While the DB has more, do it all over again
                        jobHeadWhereClause = String.Format("JobHead.Plant='{2}' AND JobHead.JobFirm = True AND JobHead.JobComplete = False AND JobHead.StartDate < '{0}' AND JobHead.TravelerLastPrinted = ? AND JobHead.JobNum > '{1}'and jobhed.prodcode='{3}' BY JobHead.JobNum", Date.Today.AddDays(DaysAhead), joDS.JobHead(joDS.JobHead.Count - 1).JobNum, _session.PlantID, "DC")
                        moreRecords = False
                        joDS = _job.GetRows(jobHeadWhereClause, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 100, 0, moreRecords)
                        UpdateJobs1(joDS, _job)
                    End While
                    '***************************************************

                Next

                '***************************************************'***************************************************'***************************************************'***************************************************

            Next

            '***************************************************'***************************************************'***************************************************'***************************************************

        Catch ex As Exception
            Dim sw As StreamWriter = New StreamWriter("C:\IT_ADMIN\error.Log", True)
            sw.WriteLine(ex.Message)
            sw.Close()
        Finally
            Console.WriteLine("---------------DONE------------" & Now() & "-------------")
            _session.Dispose()
            'Comment the next line below to get it to pause.
            Console.ReadKey()
        End Try


    End Sub

    Sub UpdateJobs1(ByRef joDS As JobEntryDataSet, ByRef _job As JobEntry)
        '*************************************************
        'FLAG JOBS FOR MASS PRINT, SET JOBHEAD.CHECKBOX19=TRUE
        '*************************************************
        Dim sw As StreamWriter = New StreamWriter("C:\IT_ADMIN\work.LOG", True)

        For Each jobHead As JobEntryDataSet.JobHeadRow In joDS.JobHead

            jobHead.TravelerReadyToPrint = True
            jobHead.CheckBox19 = True
            sw.WriteLine(Now() & "Mass print flag: " & jobHead.JobNum & " / " & jobHead.StartDate)
            Console.WriteLine(jobHead.JobNum & " / " & jobHead.StartDate & " / " & jobHead.ProdCode)

        Next
        _job.Update(joDS)
        sw.Close()
    End Sub


    Function CalcBusinessDays(ByVal DStart As Date, ByVal DEnd As Date) As Decimal

        Dim Days As Decimal = DateDiff(DateInterval.Day, DStart, DEnd)
        Dim Weeks As Integer = Days / 7
        Dim BusinessDays As Decimal = Days - (Weeks * 2)
        Return BusinessDays
        Days = Nothing
        Weeks = Nothing
        BusinessDays = Nothing

    End Function
End Module
