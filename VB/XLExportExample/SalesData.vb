Imports System
Imports System.Collections.Generic

Namespace XLExportExample
    Friend Class SalesData
        Public Sub New(ByVal state As String, ByVal actualSales As Double, ByVal targetSales As Double, ByVal profit As Double, ByVal marketShare As Double)
            Me.State = state
            Me.ActualSales = actualSales
            Me.TargetSales = targetSales
            Me.Profit = profit
            Me.MarketShare = marketShare
        End Sub

        Private privateState As String
        Public Property State() As String
            Get
                Return privateState
            End Get
            Private Set(ByVal value As String)
                privateState = value
            End Set
        End Property
        Private privateActualSales As Double
        Public Property ActualSales() As Double
            Get
                Return privateActualSales
            End Get
            Private Set(ByVal value As Double)
                privateActualSales = value
            End Set
        End Property
        Private privateTargetSales As Double
        Public Property TargetSales() As Double
            Get
                Return privateTargetSales
            End Get
            Private Set(ByVal value As Double)
                privateTargetSales = value
            End Set
        End Property
        Private privateProfit As Double
        Public Property Profit() As Double
            Get
                Return privateProfit
            End Get
            Private Set(ByVal value As Double)
                privateProfit = value
            End Set
        End Property
        Private privateMarketShare As Double
        Public Property MarketShare() As Double
            Get
                Return privateMarketShare
            End Get
            Private Set(ByVal value As Double)
                privateMarketShare = value
            End Set
        End Property
    End Class

    Friend NotInheritable Class SalesRepository

        Private Sub New()
        End Sub
        Private Shared states() As String = { "Alabama", "Arizona", "California", "Colorado", "Connecticut", "Florida", "Georgia", "Idaho", "Illinois", "Indiana", "Kentucky", "Maine", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nevada", "New Hampshire", "New Mexico", "New York", "North Carolina", "Ohio", "Oregon", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Virginia", "Washington", "Wisconsin", "Wyoming"}

        Public Shared Function GetSalesData() As List(Of SalesData)
            Dim random As New Random()
            Dim result As New List(Of SalesData)()
            For Each state As String In states
                Dim targetSales As Double = (random.NextDouble() * 500 + 40) * 1e6
                Dim actualSales As Double = targetSales * (0.9 + random.NextDouble() * 0.2)
                Dim profit As Double = actualSales * (random.NextDouble() * 0.1 - 0.03)
                If Math.Abs(profit) < 1e6 Then
                    profit = Math.Sign(profit) * 1e6
                End If
                Dim marketShare As Double = random.NextDouble() * 0.2 + 0.1
                result.Add(New SalesData(state, actualSales, targetSales, profit, marketShare))
            Next state
            Return result
        End Function
    End Class
End Namespace
