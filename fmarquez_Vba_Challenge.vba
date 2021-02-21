{\rtf1\ansi\ansicpg1252\cocoartf2513
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww35640\viewh19640\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub why()\
\
    'Define variables\
    Dim Stock_Ticker As String\
\
    Dim Yearly_Change As Double\
\
    Dim Percent_Change As Double\
\
    Dim Stock_Volume As Double\
    Stock_Volume = 0\
    \
    Dim Total_Stock_Volume As Double\
    Total_Stock_Volume = 0\
    \
    Dim Open_Price As Double\
\
    Dim Close_Price As Double\
    \
    Ticker_Counter = 2#\
    Ticker_Open_Close_Counter = 2\
    \
    'Ticker Place on Summary_Table_Row\
    Dim Summary_Table_Row As Integer\
    Summary_Table_Row = 2\
    \
    'Column Headers\
    Range("I1").Value = "Ticker"\
    Range("J1").Value = " Yearly Change"\
    Range("K1").Value = "Percent Change"\
    Range("L1").Value = "Total Stock Volume"\
    \
    'Last Row\
    Dim Last_Row As Double\
    Last_Row = Cells(Rows.Count, 1).End(xlUp).Row\
    \
    'Go through tickers\
    For i = 2 To Last_Row\
        Ticker = Cells(i, 1).Value\
        Stock_Volume = Stock_Volume + Cells(i, 7).Value\
        Open_Price = Cells(Ticker_Counter, 3)\
       'Conditional to check ticker\
         If Cells(i + 1, 1) <> Cells(i, 1).Value Then\
         Close_Price = Cells(i, 6)\
         Cells(Ticker_Counter, 9).Value = Ticker\
        Cells(Ticker_Counter, 10).Value = Close_Price - Open_Price\
                ' if statement\
                If Open_Price = 0 Then\
                    Cells(Ticker_Counter, 11).Value = Null\
                 Else\
                    Cells(Ticker_Counter, 11).Value = (Close_Price - Open_Price) / Open_Price\
                End If\
                    Cells(Ticker_Counter, 12).Value = Stock_Volume\
                Stock_Volume = 0\
                Ticker_Counter = Ticker_Counter + 1\
            End If\
\
     Next i\
\
End Sub\
}