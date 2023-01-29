# TOUHOKU (PCA) WEIGHT SOLUTION
Solution to help 西山 team of エマール group transfer data faster for 東北 (プレキャス) 重量 type normal from 文化シャッター partner.

## MASK
<p align="center">
<img src="https://raw.githubusercontent.com/Tynab/Touhoku-Pca-Weight/main/pic/0.png"></img>
</p>

## CODE DEMO
```vb
''' <summary>
''' 運賃 (2トン車).
''' </summary>
''' <param name="xlApp">Excel Application.</param>
''' <param name="choosen">Selection.</param>
Friend Sub Fare(xlApp As Application, choosen As Double)
    If choosen = 1 Then
        DctVal(xlApp, "BA150", choosen)
    End If
End Sub
```

### PACKAGES
<img src="https://raw.githubusercontent.com/Tynab/Touhoku-Pca-Weight/main/pic/1.png" align="left" width="3%" height="3%"></img>
<div style="display:flex;">

- Microsoft.Office.Interop.Excel » 15.0.4795.1001

</div>