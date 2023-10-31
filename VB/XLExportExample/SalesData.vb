Imports System
Imports System.Collections.Generic

Namespace XLExportExample

    Friend Class SalesData

        Private _State As String, _ActualSales As Double, _TargetSales As Double, _Profit As Double, _MarketShare As Double

        Public Sub New(ByVal state As String, ByVal actualSales As Double, ByVal targetSales As Double, ByVal profit As Double, ByVal marketShare As Double)
            Me.State = state
            Me.ActualSales = actualSales
            Me.TargetSales = targetSales
            Me.Profit = profit
            Me.MarketShare = marketShare
        End Sub

        Public Property State As String
            Get
                Return _State
            End Get

            Private Set(ByVal value As String)
                _State = value
            End Set
        End Property

        Public Property ActualSales As Double
            Get
                Return _ActualSales
            End Get

            Private Set(ByVal value As Double)
                _ActualSales = value
            End Set
        End Property

        Public Property TargetSales As Double
            Get
                Return _TargetSales
            End Get

            Private Set(ByVal value As Double)
                _TargetSales = value
            End Set
        End Property

        Public Property Profit As Double
            Get
                Return _Profit
            End Get

            Private Set(ByVal value As Double)
                _Profit = value
            End Set
        End Property

        Public Property MarketShare As Double
            Get
                Return _MarketShare
            End Get

            Private Set(ByVal value As Double)
                _MarketShare = value
            End Set
        End Property
    End Class

    Friend Module SalesRepository

        Private states As String() = New String() {"Alabama", "Arizona", "California", "Colorado", "Connecticut", "Florida", "Georgia", "Idaho", "Illinois", "Indiana", "Kentucky", "Maine", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nevada", "New Hampshire", "New Mexico", "New York", "North Carolina", "Ohio", "Oregon", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Virginia", "Washington", "Wisconsin", "Wyoming"}

        Public Function GetSalesData() As List(Of SalesData)
            Dim random As Random = New Random()
            Dim result As List(Of SalesData) = New List(Of SalesData)()
            For Each state As String In states
                Dim targetSales As Double =(random.NextDouble() * 500 + 40) * 1e6
                Dim actualSales As Double = targetSales * (0.9 + random.NextDouble() * 0.2)
                Dim profit As Double = actualSales * (random.NextDouble() * 0.1 - 0.03)
                If Math.Abs(profit) < 1e6 Then profit = Math.Sign(profit) * 1e6
                Dim marketShare As Double = random.NextDouble() * 0.2 + 0.1
                result.Add(New SalesData(state, actualSales, targetSales, profit, marketShare))
            Next

            Return result
        End Function
    End Module
End Namespace
