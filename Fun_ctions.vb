Option Explicit On
Option Strict On
Imports System.IO
Imports System.IO.Compression
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Text.RegularExpressions
Imports System.Data.OleDb
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports org.apache.pdfbox.pdmodel
Imports org.apache.pdfbox.util
Imports System.Reflection
Imports System.Globalization
Public Module Functions
#Region " GENERAL DECLARATIONS "
    Public ReadOnly Segoe As New Font("Segoe UI", 9)
    Public ReadOnly Gothic As New Font("Century Gothic", 9)
    Public ReadOnly Desktop As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
    Public ReadOnly MyDocuments As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
    Public ReadOnly InvariantCulture As CultureInfo = New CultureInfo("en-US")
    Public ReadOnly Property WorkingArea As Rectangle = Screen.GetWorkingArea(New Point(0, 0))
#End Region
#Region " REGEX DECLARATIONS "
    'http://ascii-table.com/ansi-codes.php
    Public Const phpRegexRecursiveParenthesis As String = "\(((?>[^()]+)|(?R))*\)"
    Public Const BlackOut As String = "■"
    Public Const Delimiter As String = "§"
    Public Const Heart As String = "♥"
    Public Const NonCharacter As String = "©"
    Public Const EmDash As String = "—"
    Public Const Space As String = Chr(32)
    Public Const NewLine As String = Chr(10)
    Public Const ListSeparator As String = vbNewLine & "❶○○○○○○○○○○" & vbNewLine
    Public Const AlphaNumericPattern As String = "([A-Z]+\d+|\d+[A-Z]+)\w*"
    Public Const NumberPattern As String = "[0-9]{1,3}([,][0-9]{1,3}){0,4}[.][0-9]{2}"
    Public Const CommentPattern As String = "--[^\r\n]{1,}(?=\r|\n|$)"
    Public Const FilePattern As String = "^[A-Z]:(\\[^\/\\:*<>|]{1,}){1,}\.(txt|gif|pdf|xls|doc)"
    Public Const SelectPattern As String = "SELECT[^■]{1,}?(?=FROM)"
    Public Const ObjectPattern As String = "([A-Z0-9!%{}^~_@#$]{1,}([.][A-Z0-9!%{}^~_@#$]{1,}){0,2})"     'DataSource.Owner.Name
    Public Const FieldPattern As String = "[\s]{1,}\([A-Z0-9!%{}^~_@#$]{1,}(,[\s]{0,}[A-Z0-9!%{}^~_@#$]{1,}){0,}\)"
    Public Const FromJoinCommaPattern As String = "(?<=FROM |JOIN )[\s]{0,}[A-Z0-9!%{}^~_@#$♥]{1,}([.][A-Z0-9!%{}^~_@#$♥]{1,}){0,2}([\s]{1,}[A-Z0-9!%{}^~_@#$]{1,}){0,1}|(?<=,)[\s]{0,}[A-Z0-9!%{}^~_@#$♥]{1,}([.][A-Z0-9!%{}^~_@#$♥]{1,}){0,2}([\s]{1,}[A-Z0-9!%{}^~_@#$]{1,}){0,1}"
#End Region
#Region " IMAGE STRING DECALARATIONS "
    Friend Const EyeString As String = "iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjAuMTZEaa/1AAACYElEQVQ4T6WUv24TQRDGLwYCNCAkQEFI2dm9vfg4yZUbS4i49QNgkBANjUusPIHfgIYoKYGKhqShQKJJmQoBAiEEDYI3ICilw/ft7ZzXxEFIRNrs7Ow3v5s/vssmk0kry7Kl4XB4Ktot2vTxrDb31E7jGMNz9GVL3W73DC/6/f5ptbnzTHs0GgVfai/SBqAK/xem/lDy32C9Xu98ae0a/sq+MedO0tJm6Sy5tUhQSLGeG3mBdejFHmHnOsxFdpxzN/+EcedqSlaB9/4CAE8QPCXEi7zxRjaxHudi39ZgO/XGPu06dzGFMb4ZCg9FUVyH+F3M5qe3Frr5nsF3G3cH4WFGPlbOrdLP9pDV9BA9uoxyPkXYUQrzxj1ApuOqqlbCGdBZG+xn+qnlCiUHkdhXDQxl0s+suPCgRwQ4I9/RkivMBr5QfgS/JiMAmeaa5HdnMMtSNuvM7AZ8L724sWZUiDxkIIBb6gtL8juhZF4i8J7Cwo4BMDNvzEZ4gPqxz4B2S2FBY+39kCH/hWCxeyrgNOkvy/Iazj80CP36xl7jblmHF/wi+4PB4Gzz6rEnlTErEHypAyHEUFg2fnNX65JlTBi1vFMY+8pJMymw5l+9drstEH6I0IM46WXV0I6wXxH29YZzBWG8J+PYq9fpdC6hp88QNK3BKE1kG23Yhu89lpb5nBlrXJMhS06dhNMurL2Vr8qOZkNQtHcxmPU0CY2b+9qkMBVy8eOQ5zl+fj6nTd8imPqZYfg4LIJxAGqr/yTYsVcvDVKB2v8CU20AKpTZaiCf1oiiHXoEjWqTuHg/af0Gh2dvChU5Q8wAAAAASUVORK5CYII="
    Friend Const ClearTextString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjAuMjHxIGmVAAAB5ElEQVQ4T41TTUtVURRdzyLNSU1MCSppEChkkKbkrK9JOJBAhCai9u45R1DoD/T+QrOgQTl2IkiI75xzuzpIkSQcaCmlDSoJHEhBBEW91vnQR0+eumDDvXvttde+Z58L9E0cw5A+i4GsDkfFSNaEvL1K7Qkgb65B6DeQ5jkScyqWVMewbmHdIoTZRqJvAcregLQ/mPjNeHJAkxyEbYXUKzT7y+dPeGAuAclUPSd45JtI84sNnkIWz0RRGcq2sW4+iM0mJ7+NQqEmkO77k3SQ5E+Sf+jyAoPTDYEkkqyD4657sTRbdO9EoRTFuyhkx0mOkfzmmwg9jmHbCDFzhdO9Dc52HeplO6tzQVSJ0elaOrlJdsIkRjOW2aDE9zWotLs8djW4tSp9j8LvXuiD36xmW4BSFedKJOY6Hd+XG/BMRNYc2UMgTQ/jY3ReZXyOjV7786gOjifsHRZ+obtb1ZJ3lcWbbLLFcBt4BWUuRkEFRHqXBR+8m7ALyOvLPu8OTdle5jbiVLyFXON/kOl9El+D2CxgJL0QmYASp3PrE3b3XN5xWx2BTMx5ildJurXNQc2cC8Q+sIm79ntTPsPD+ZO8hZOnmXzMxHi42wcih3yxi6bux+v3l8/D/Q/JUn14OQKcsxcD/wAb8/HwyOWubQAAAABJRU5ErkJggg=="
    Friend Const DropString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjAuMjHxIGmVAAAA3klEQVQ4T2MYBVCQvkeLIWNvLEP9fhaoCH4Qv18AyoKC9N0PGDL2/AfSyxlCVzFDRTEBSC5j9wyI2j11UFEgSN+zFCwIwQsZ6uuZoDIIELqKDah5NULd3ulQGSDI3cYOFNyKkNwzi4HhPyNUloGh8Bgninz67m0MaWe4oLJQEL+fA+iS3UiKpoDFk47wAtkH4OIZe1YCNbOC5TAAyNT0PUiK904GOvsEnJ++Zy7eMAKDpA28QMXHEIbANU9C8RZekLabH6jhFELz7laoDAkgcZsoUPMSIC6AigxPwMAAABk/eh6Y0kgKAAAAAElFTkSuQmCC"

    Friend Const Edit_String As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABl0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC4yMfEgaZUAAAFvSURBVDhPlZMrSARRFIbXRREfaFBUENuIsDM7D5YZ1uIUgwYtilksJvMiIgajoBg0iGnTBsFiEKNBRCzCVmERTMLCGlaDr+/oWVD0jvrDx73n3P+c+xgm9R85jtPved627/txLpdr0fTfFQTBFA2eoQYlmvTq0u9i12XXdVcpnIcT4gco63KyKIigDo9QgQLFkzScUItZHLObgnM4oKCPwnXmNcYjWVObWZjnKKyDIzFjhxwdjvP5fNu7yaQoirpocIV5TVPykNPEVYg1ZRamLShnMpkBTTXR8JJTbMr8I2UQhXi9O3Zc0lSauADXMKg5szDtw5llWa0SZ7PZEeJbWCFMS84ouR/Ge3Yf1ZQ03NDrdGrqZ8Vx3Iz5Ap4oKIZh2MNc7lOl4ZjakoX5lOJXgbkc+4b5Id+9XS3JomCn0UB54eUDXf4qFmZ4oNkGHNOCRYrkp6kwlmBc7d+FaRfD3idiGKbxgm3bQ8m/bSr1BlLvcgz+uCnlAAAAAElFTkSuQmCC"
    Friend Const Add_String As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABl0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC4yMfEgaZUAAAF+SURBVDhPY8AFzMQ3lWqq1c3XUe6ebyFwTBcqTDwwFd20S0u9/r+ucvd/a8FTnlBh4oGp2OadMANsBU54Q4WJB2ZiW3bAXSBwyhcqTDwwE9u8HWJA138rgdP+UGHiAdAFW0EG6AANsOE/EwAVxgQWorvjzIR3zDQT3TnLXHTnbCCeayayc56R5LLHYANU2/6biW09YCa6fYmp6LZlQHo5EK8wF92x0kJkuxGDplrDfC31hv+YuJ4g1tCo9WbQVmmcr63W/B+MVVsQWK0JqrAByG8FuwSC2//rqECxars3gx3vVS8bnstltryXS+x4rxTZ8F4stOO9nG8muvkSyACQQhuBU3NseS6l2/NcTLXluZhix3shyZb3YoK90AkZaEhgAlPRLeshBnT8t+G7GAkVJh4AA24tzABbvkvRUGHigZnYtjVwF/BeiIUKEw/MRbevgoUB0L/xUGHigZnotpXwQOS9lAgVJh6Yi+1wBWbnTG2V1kw7nrOaUGE0wMAAABm0uZSULDJ3AAAAAElFTkSuQmCC"
    Friend Const Remove_String As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABl0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC4yMfEgaZUAAAHUSURBVDhPY4CBcsEz/EkM13mhXJwgg/uCWC7DNnYoFwJCGf4zn5SJXX1eNnJ3DMM/bqgwBkjnvCl9W8Hj0kbxqkaoEANDL98OodOysWv+acr//68p9/+afMDGmTwzRaDScLBavEP7gZLDDZCaXxoq/w/I5NTXM1xhY6jn3CV1X9HpFkgChP9oKP65KBu+F9kluZwXZR4qOd34qyn/D6bunGz4mniG+xxgBfX8h5UeKdmf+6upAFdwVS5gE8glyDaD8C8N5T9XFAI3ZzI8FARrhoGFIj1qjxQd7sAUglxyTjZsP7rNV+QDN7XzL0XVDAPYXALDIJuvAjXHM7wXgCrHDtBdAsN4bUYGID8/VHK8iW7ANXk/cJhAlWEHoNAGBRiyn2H4t4biX2DI7w1leMkDVY4KIDY7oYT2VYWALehhck3ef+NC3oXCUG0QAEph6DaDAgzkZ1CYPMCIHSSX1HNuk7mn6HoVpgAcz2ihDYkdBxSXnJOLXJnGcIaLoVT4CO9RmYSNMAlQIsEW2sgu+auh+H+HRAkwKe9nAUt6MPxjPy0bveSCXOh6jBSGBEAuuangfWSneHEDA8N/RqgwBHSL7+SuBzmJAKjnXy+QxjCTFcJjYAAACcw1wQVwmDAAAAAASUVORK5CYII="

    Friend Const SortString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjAuMjHxIGmVAAABRklEQVQ4T+2TsUrDUBSGm9ykMUSJEGJAhDhkimDRLAUtwVcQUix0EJw6+AwZ3Vyd1cWH8AUsODi6CBYfQLqooNb/h3tqokPr7vAluec/5z8nyb2NsizNoiiUkCSJk+e59TMWhuFiEARLURR5Vb3BCwPEdd1VwzBOLMvaTdO0yVgcxwumaR4hfgEuwSny1sRkauA4zjrEa/AGhrZtt7Iss9kV63Mw0TxC26oZ6M4sftdJn+CWnbQBu1cNtqcGHBUjDiDcAxZK4gM4m2kgrwCTPsRXScT6kB/vLwZdiFWDHqf7N0ibvu8vY30lcfCEjdae2wDPsrnEgL96BJMd1s00wLibWN+BD609QzvmmahNoJTah/iik2jQpQG3M7p1ELsBYxbzu/x6Bc/zViAegD7v3N6MEz3JBprsyWmsGQjsJlTjghR9U6gvKYe7gZJsyxAAAAAASUVORK5CYII="
    Friend Const CheckedString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjAuMjHxIGmVAAAB1klEQVQ4T2OgK/j37x/7nqMnDi/fubUIKkQauPb09vSZOlf+z1a78X/pro1VUGEIcHBwyHN0dIzChWfMnrGx2H7K/07xXf97RPf/79A/8H3pmaUiUO0MDPb29hFQJgb4//+/yMLJu95uYPrzfyPj3/8dErv+Rzm3F0KlIQCXAUDNzBt279u1iPfl/01AHsgAH8+kHQz/GRihSiAAlwEXHl3um6F1GawZhPuVTr0Gqg2FSiNAampq3snHZ+YAQ1oKKsTw7sdHv+lRh+Ca5wk/+pdaX+OHYcCbf2/4ctPrrkxVuvx/894Da0BiQKdzzO3d8QbkZJDmtSw//pfmTpscGhrKjGIAUCFTQ+Ginc0ym8EKZ0nc+RNfUxnR1Du7eAn3G7jtzUHrL4MMxTAABKo7J9dVaiyGK14g9vj/TOXrf2H8KUZnvs07tkQXpBarASCQm1m5ZoHIE7ghMLxA9Mn/aauXZUOV4TYAmFjiJk5ZuGaW0K3/y9negvESjhf/G6tm7nV3d5e0tbUFYycnJ2msBtjZ2akCU6N9y/TOBXXaa/5PE7jwvzS5946bm5sjSBwZ29jYiEK1YQJgQDHOWLWiqdF/3bcNV7cZQoVJB/ff31eAMvEABgYAPDoU9wQut5cAAAAASUVORK5CYII="
    Friend Const CheckString As String = "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAYAAAAfSC3RAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABh0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMS40E0BoxAAAAN9JREFUOE+N0rsOAUEYhuF1SBwShUaicRN60aqIkkgU7kChUqJW0LkCpxAKd+BQaEm0DheglXg/WYK1MX/yZHcm+83OTH6LqmOBuYsZxvZzYs/1YI0wRRkFF0WU7Ocae1hDNBHUwKAG2OlFwQbcgiF0kYcHRsEAarhCAY2NgimcsEJCE9TPoB8xqDTe4oykJuxyBHWWDjZIow9tsQIt+CxHMIwcDrjgBi0UwXv93KoPWRyxRBTf5Xo5uvIM4o+Rsz6CLWirXgOvoFpOLaS2U0/+o2M8Wq4K/VUrmeBbq30HmqxP1SI+lSYAAAAASUVORK5CYII="
    Friend Const UnCheckString As String = "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAYAAAAfSC3RAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAYdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjEuNBNAaMQAAABvSURBVDhPYwCCWiDeCcRbceAtQLweSm+Eis0FYoZ1UIEEII7CgaOBOBZKnwLim0DMsBaIW4CYA8QhAqwB4hsgxqhGTDBUNbYBMRcQMxGB4RpBSQ6UhDYBMShNEsJPgRic5IqBGGQryCRiMFAtw0QAbhE+zDCrvcQAAAAASUVORK5CYII="

    Friend Const DataTableString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAmJJREFUOE99Ul1PGkEU3T60Puiv8iPR+OqrH0++m0ir0TQxJjWR4EdrCBgBG4wQSyEiQmhgXUAqKwtbwGRZ2Fl2MaAoghWILCxMh5AaNU3v070z98w55955AyHE/gbP85Ikoardbv9+qLx914NupXq9r7ev1YLtNuzt7cHQ0VOwLPuUcwCgHPUxLF+twrs7mM9Dkoy9ACSTyVcAVLIpvlbrAK6uYChEY7IMRbFB01W/v3R8nLNaBbM5ZTQm9HrSbA4TRIJNgVpNLpXkQkEOhaLY5aWcycjRaP30tOL1llyugt2es9nEvT1ma+tselqr1pgZBsTjgKYBQQSxTKbJcS3E4PPdOByi1crr9b+USu/8vGVy8vPg4AfDrv3hAd7ewlwOBoMUxvMNBAiHKy5X3mRKoYfVanJl5cfc3LeJic2Bgfe7Xx3lMry+htksDAQoTBAaglCnqNLRkWAwxLVaSqUilpYcCoV5aurL6OhH1ZoxFgORCDg/BzgewACQugxeb8FmQ47T29uR1dWOpC6DZvswn2+KYhOAps93jjw8ptPSPyWNj28MDy+oNYeiCDkOMgzEcRLjuEdBaFFU0e3OmkyMTkevr/uXl51dSUNDcxub39EO0YYuLqDHQyJJtf+YRlNaVeoIgsTxkMdzRlE0YqgJgkSSRYulox6ZVirxxUXbzMweYhgZWTw5oZ9/HzTWKsvKkUjF7y/a7Zf7+yyCra35FhasyHR/vyIaTb8AyHKbpks4fuN05p7vYXbWNDb2aWfH+by787W79f19PZG4JgjO4biwWKiDg59uN1UuV151o/IPPuNL2ItzNKQAAAAASUVORK5CYII="
    Friend Const StarString As String = "iVBORw0KGgoAAAANSUhEUgAAAAwAAAAMCAIAAADZF8uwAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAXtJREFUKFNtkU1LAlEUhkcQbFt/xT/gylX+AJcRXDDCRS7CCCwCFzXgJsiLiYQwOWUZVx2/onGERO9cJ7OoRWbgJ+aIEgWGKF01AqGXszjvy3M4cI6m3W6rqqrVajUaDTOv8Xis0y0sLS0yGONyuVyv15vzarVatVoNYzmfx0wul6tWq43/RCdlWc5ms0wmk8HKg/z4Oqt3tacUe382mpQkKcOkUhIMlZyowuF+5K7T//x+fmleK90g6frkijd8k06nGf9FwmjjDDYuFH+jxGg0GgyGjUbnHD25OLLBomhUZLxngh5APfCbtkmCNIbDYfPjy+q7N1gIDfXAwSH0Cxnt4tZJ4Vaph3FVrrWOkmXgJEYg6gHLocQEMgK4yoqb3oLrtLR2rOwGlH1/EUBi3qMQ5NB0nWkHAihaIHFCAqdFG2oBKy7PoMBlxGxjgSO4DsOsG3k8AoTowI2oBfbgivUwkpImF+f5K54XeEEICUIsFhemDbU0RChOH/EDcKchcY4euAgAAAAASUVORK5CYII="

    Friend Const BookClosed As String = "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAIAAACQKrqGAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAALDgAACw4BQL7hQQAAAddJREFUKFNtkd9P2lAcxYsPvvnv6njBYIAhEKVhxkHQoREFtOWHcYEZQKoIpSkt0dKglh+9GAwphlbBUSzsEhyyZOfp/vice3LPVzcej5E5jUZjRVF1OmRpaXFhQTd/hUB0puG7JtTkZArEYkKZl7rd3yNo/atPVNNGQlWOnwMsLJz9bMYTtdRFg2XbqqpN4Q8UcvWGEjmtneBVPAwIQk6nnn8l2jguRqL3HCdBwwSd5FZlyIUj9dBxPRGXotFW6Bj4/VWvl9/YYEzmPGgqyCS3Jp8nAXzvwC9gWDMYaPh8925PGUVLViv1ZSWztpZqtWSk03kLYRXf3u3Oj7sd7517m99ycc7NksPBWCyUXp8xmZNZku313xBRlLe+k55dat2RW7dR1q+UxUyajKRxNafXE1Zb+qZye3RSBEBGRPDs2SXyDB/AaLMFZuWMxvyqIQs5myPNcDeQc7kKktRD+v3B1fXDvr+QpfnLAv9tmzQYrlaWCTvkylwQKzqd5NPT60dZw6HWbr8eBpjwGctWhCOM3kQvqBIbwot2+/XLy+CfXuFGVd9zefEwyNDcA+g8YjEWRckZ9zmCqQ8W1xC7mUshgDFuDz3N/c9gZ0e93qD5qMB/zHNw/QdY3clc1dADtgAAAABJRU5ErkJggg=="
    Friend Const BookOpen As String = "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAIAAACQKrqGAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAALDgAACw4BQL7hQQAAAeJJREFUKFN10V9v0lAYBnCWJV5544fxM3jntYlGs5joxbIlOhNHNEsVNS4zkexPtiUsm8ZZNlmxIRtsrJSuDih1qXQUgrUtpS3dpCtDkI5hwRYSuJg+d+c5vzfnTc5Au912dKJVGlcuXxocHOge/xGLWuGkX/t0KZuvNM7NbnMxDqtS1foSmNHKxhf6+IdctRrjrElli0GcRZJib8am6bQ25Sb1spHmNEY4zXA6tENHknm9Ykx7v9VqzfPOUzalqKOXr/d1vc5JJT/KfsZ4SdXOGk3r6tUi5t8QcOyoXm/alPyqAC7MogVFW92iITTXp/NBglTmZ9Nq8bdNiaQ8AaBdurSRWN9levTFXIA8KE6/TRWVmk3xGD/+NHTSoQsg5t2me/TZjM+iU5MHslR1tFqtIEKNjPkSRMGic6sRMJQSlZJyfBqn2GVfNE5Iz4GYKFYs2mZy+VGnB3CF0T125kN4PXwYJRj3SugdhB/mJAwXnOOoIJTtBf6YJisow48X3bN7bzybd50LgBvywgksnksxcmArO/pgk+NPbNoNX1DHJlau33JdvTZ047bn5tDynXvv7w+Dj5zwwydrvPizT03T/M7La3DkI7QD+ndBGPHCyKdANIiQMTJTrRl9+r+v7/V/AT5wyfCHirK9AAAAAElFTkSuQmCC"

    Friend Const ArrowExpanded As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAM9JREFUOE/N0kkKg0AQBdCf+x/KlaALxXkeF04o4g0qXU1sDPQqBhLhI72o17+gH0SEWx8Dd3JrWLa/c/t3Ac/ziOM4Dtm2LWNZFpmmKWMYhsq1tVphmiYw0Pc9tW1LVVVRnueUpilFUURBEEjAdd23tdVhWRZckaZpqCxLiSRJIocFwpfogW3bwMg4juDqXdfRifCwQCCawPd9PbDvO9Z1VQgPMcJ/0QRZloGRMAz1wHEcOJF5njEMA14I6rpGURQSieNYD3z6Hv7oIf1shSf3G9UMQ+Vu/QAAAABJRU5ErkJggg=="
    Friend Const ArrowCollapsed As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAANFJREFUOE+l00kKhDAQheHq+x/KlaALxXkeF04o4g1ep0J3Y0OEiILoIvnyK9QLAD26GHhyKzc7joNpmmhZFtq2jfZ9p+M4lGsvgTOyrqtEVKWXQN/3YGQcR1nCiDZg2zbatsUZmef5HlBVFZqmQdd1smQYBn3AsizkeY6yLP8Q7U8wTRNpmv4QLhAl+gUMRFGEJEnA76KE6rrWBwzDQBAE4KdAKMsyKoriHvBBSJTQF9H+B7zZdV3yPI9836cwDCmOY/2CO7PxaJDkJN85TbX2Db5d1YfJcQ3TAAAAAElFTkSuQmCC"

    Friend Const LightOn As String = "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAIAAACQKrqGAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAUtJREFUKFNj/P//PwMK+P/r6ytGRmZWLhFUcQYGoFI4+PryxMfbk388X/rj2ZIPt6d8e3MRWRah9MurC9+er/375fi/T3v/ftr99+vJr49Xfnt3C64aofTtrSV/Ph76/bju9/Npv59P+v245s+no29vLsWi9M3VGT/uVPy4lvTzQdfPe80/Lkf+vFvz+soMLEpfnu/4diHy683Wn6/W/3y96euNpm8Xol+c68ai9NnF+Z9Oer/cYfJ6n8Orfc6vdpp+Oun//PIyLEp///hyb1fS620Gd1ca3Ftl8Gab4f29eX9+/8SiFCj0/fPrh9tib6/2vLnU6N6mkF/fP2EPLIjoo2OTrq+JuDLf5NHRCcjqQDGFxr91fPn2/oC1zTa3Tm8koPTIkSN7du8qKS4+deoUAaWXL19uamoqKSm5ffs2AaVA6b9ggKYOyAUAkObu3QMxkwMAAAAASUVORK5CYII="
    Friend Const LightOff As String = "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAIAAACQKrqGAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAARZJREFUKFOFkM1qg1AQhdOdT5uF6QM0O9MHaHYppFtXulJBgxupeiVWEfzb+IMUBK1YQRHtoKCmCJ7FZZj57rlzz1Pf97tHNU3TdR2GYf/6O0AnpWmqKIqu64ZhqKoahuFyOqN5ngNUVdXvoLIsPc+L43iiZxSciqIA459BUMBlTdNWULBMkiSKou9BUARBgBCCvUd6dgUD3/cdx4F3Qa7r2rYtSdKKq2VZ9/sdbOA0TVOWZag1hFbQtm1FURQEgaZpiqJ4nuc4Dn65gkKrrmuAPq7X89uZJMksy9bDGrvyp/x+uRAEwTDMknv41jj40vXjy3G/38MmGyhC6k24PR8OLMtuoJDX6+mE4zjEvIHCGDKfYl/Sf9M5/Uxpz2tBAAAAAElFTkSuQmCC"

    Friend Const DefaultExpanded As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAO9JREFUOE+lk9sKAVEUhnkoD+E5vJJzuHEshzs1DiGUQ0RRigsKJceRcWZ+tppxmjXKTK2bqe/791p7Lz0AnaaPCbSUDI8mK3iiBbgjebjCOThCGdiDKVh9SZi9nFylWvue9wyVBZ5YEaIIXK4ijqcrtvsLeOGMGX/EeH7AYLJDdyjAYDQpC9yRwk+43d/QAnZstWQGN3prWuC890wdW4IrHZ4W2AJpuWfW52cxuNha0gKLP/E1sNdkBmebC1pg9nFv01aCU/W5ukC6KgrmqjN1AbtnNThentIC9sKUhvf5j3yJ/+6DpkV6bPK/yRJ3A/PE7e2oP8DgAAAAAElFTkSuQmCCAPjCzMoz/hO+xEPvwdYhbS75UGdNtwLNm+LI5h1FwAAAAABJRU5ErkJggg=="
    Friend Const DefaultCollapsed As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAARpJREFUOE+lk9tKAmEUhfWhfIiew1eyUkkvPCaNdZdHbPCACIqCklZERQZCZhNZeZylS5jRhvlHcAb2zcD61tqLfzsBOGx9BNgZXdwffCKYLCEgFXF2IcN3XoA3nsNJJAtPOK1Ptd5Z+21NdUDwsgxVBRZLFdPZEj9/CyjjOd6VKd6GEzwPfnH/OobryG0OCEilvWK58SIGMLbRmW6ac+fpG1fynRjgX+9sjE0AY1PcfPiCVLAAnMby+s4UGqfWVZDI98QJjqMZ08LoTHG5PUI82xUDPJH0v7YZmyk08U3rA9HMHsBuYbvOFOcaQ4RTt9YJWFix2Ueq8rgpjDszNp0pDl1bAPjCzMoz/hO+xEPvwdYhbS75UGdNtwLNm+LI5h1FwAAAAABJRU5ErkJggg=="

    Friend Const ChevronCollapsed As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAKCAYAAAC9vt6cAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjAuMjHxIGmVAAABRklEQVQoU2PABz5//uz06dMneyiXNHD8+PFyGxubf2ZmZn+PHj1aDhUmDgA1FxkbG/8HMsFYR0fn/5EjR4rAkoTAiRMnckxMTP4BmXADQBhoyL9jx47lANm4AVBzuqmpKVgzIyPjf3Z2djAGsUFienp6IEPSwYrRAVBzEsi/QOZ/Jiam/83NzQf+/fs3G4QnTJhwnJWVFWyIvr4+KEySwJpg4PXr17HJyclwzSUlJRv////PCpFlYADaytnU1LSbhYUFbEh0dPSfL1++xIIl7969K/cHCCZPnvxPQEDgX2lp6RagZjawJBI4c+YMF9CQ/YKCgv96enr+/f379+eVK1ckGG7evCkN5PwCavoPpHc/evSIE6oHA+zfv5/n+/fvR6Bqf9y5c0cMLPHx40err1+/ZuHTDANAW3mACSz7w4cPZgwMDAwA7Fq34WL8tRIAAAAASUVORK5CYII="
    Friend Const ChevronExpanded As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAKCAYAAAC9vt6cAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABh0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMS42/U4J6AAAAUNJREFUKFNjAIGPHz9aff36NevRo0ecYAE84MqVKzyfP3/O/vDhgxlY4ObNm9J///799R8IgPRufIbs37+f5/v370egan/cuXNHjOHu3btyf4Bg8uTJ/wQEBP6VlpZuAcqzQfXAwZkzZ7iampr2CwoK/uvp6fkHNOAn0DUSYMnXr1/HJicn/wUy/zMxMf0vKSnZCDSEFSwJBMeOHeMEat7NwsLyH8j9Hx0d/efLly+xEFkoOHHiRJKZmRnckObm5gP//v2bDcITJkw4zsrKCtasr6//9+jRo0lgTegAaEi6qanpPyDzPyMj4392dnYwBrFBYnp6ev+ArkkHK8YFgIbkmJiYgA1Bxjo6OiDNOUA2YXD8+PEiY2NjZM3/jxw5UgSWJBYADSm3sbH5BwoXoJ/LocKkAWCCcfr06ZM9lIsFMDAAABo0t+GfVFaJAAAAAElFTkSuQmCC"

    Friend Const PdfString As String = "iVBORw0KGgoAAAANSUhEUgAAABIAAAASCAYAAABWzo5XAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAYdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjEuNWRHWFIAAAH7SURBVDhPrZNBUhNBGIXnCIgX0BtwA72AVR7AQ7h2pQaIxAABN5ZEMkkqsbKOGzcWJRAVFdBsoyKo4BAksDGZzPh8r2s6zsSKWpZ/1VeT7pr+8vrvHgfAfyExqFarqFQqKJfLKJVKcN0iCoUC7i8vI5/P497SEorFEl9NSkRiIEkYhgMCEfxEf9RqtVBwXb7+G5GSSOD3+/D9Pnq+j17PR1d0e0bkeZ6RKWV8bUKk7UhUr9d/QbJms4lGo4FarYaFxTtcMkKknphENo2eNhH5xlSiHwTILSxyyZBo1TmDN+Q12SZbEZvkFXkZ8YJskOcRz8gjZ8wIjUiSjnMWx+QrOSJtcuiMwyNfyAHZJ5/JJ/KR7JEG1w5ESiKJLX9lHe2x89EI+N45wenVa0Ziq7uyht1hkbajJKrOxAWEO7s4uXzFjJWkzTnz+9yEeSqJJB/Ielykfmg7qq77wDy9KJHdTshUhxcvmTmlOc3dxQ7n1+I9kkg9UR3zZUmURCXJEdMFTKkkKiWR5D3XrcZFOh01VhVvrEoCf7uJNmXajspK3g2LdMR/Oh3bk7jkLXkSFz3mYNQ9eUp0MkKNVU+UQkjykAxEFl178539BTdSk0ZgSYh07edzOczOzSM7O4dMNouZzG2kZzKYTt/C5HQaqakpI7l+M8UlI0T/DpwfUyqMa1e21YsAAAAASUVORK5CYII="
    Friend Const BlockString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAALEQAACxEBf2RfkQAAAB10RVh0Q29tbWVudABDcmVhdGVkIHdpdGggVGhlIEdJTVDvZCVuAAAAGHRFWHRTb3VyY2UASW50cmlndWUgSWNvbiBTZXSuJ6E/AAAAGHRFWHRTb2Z0d2FyZQBwYWludC5uZXQgNC4xLjb9TgnoAAAA10lEQVQ4T6WSsRHCMBAEPyQgcEgZFOCEiOJcCkU4IHQBFEBEwJgZChB3/L+wLInBJtiR/qSTTy9LCOEviuISclGkAR0YwGgj6ybbC1JBpAXXt5xzA22yX1eimV92cw94GLUDeADqPCRJEieYMialfqJtwdn0p41dXGc12cy7UtKYqfkCjjYfoodVLLRhlBh7bt6BjdVj7QBPwDunZl3fm1ZN4D242/gx6/rJ9GoPGJ1dpsyG8c6MzS+7ma9UeQWiT+eHzKH5y3/gaJKVf+IKiuISiuLvBHkBB+NzX3/RhhoAAAAASUVORK5CYII="
#End Region
    Public Function SameImage(Image1 As Image, Image2 As Image) As Boolean
        Return ImageToBase64(Image1, Imaging.ImageFormat.Bmp) = ImageToBase64(Image2, Imaging.ImageFormat.Bmp)
    End Function
    Public Function ImageToBase64(image As Image, Optional ImageFormat As Imaging.ImageFormat = Nothing) As String

        If image Is Nothing Then
            Return String.Empty
        Else
            If ImageFormat Is Nothing Then ImageFormat = Imaging.ImageFormat.Bmp
            Dim base64String As String
            Using ms As New MemoryStream()
                image.Save(ms, ImageFormat)
                Dim imageBytes As Byte() = ms.ToArray()
                base64String = Convert.ToBase64String(imageBytes)
            End Using
            Return base64String
        End If

    End Function
    Public Function Base64ToImage(ImageString As String, Optional MakeTransparent As Boolean = False) As Image

        ImageString = Split(ImageString, ",").Last
        Dim b() As Byte = Convert.FromBase64String(ImageString)
        Dim Image As Image
        Try
            Using MemoryStream As New MemoryStream()
                MemoryStream.Position = 0
                MemoryStream.Write(b, 0, b.Length)
                Image = Image.FromStream(MemoryStream)
                If MakeTransparent Then
                    Dim Bmp As New Bitmap(Image)
                    Bmp.MakeTransparent(Bmp.GetPixel(0, 0))
                    Image = Bmp
                End If
            End Using
            Return Image
        Finally
        End Try

    End Function
    Public Function StringToBitmap(Text As String, TextFont As Font) As Image

        Dim TextSize As Size = TextRenderer.MeasureText(Text, TextFont)

        Dim Flag As Bitmap = New Bitmap(TextSize.Width + 6, TextSize.Height + 3)
        Using FlagGraphics As Graphics = Graphics.FromImage(Flag)
            With FlagGraphics
                .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                .InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                .PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
                .FillRectangle(Brushes.WhiteSmoke, 0, 0, Flag.Width, Flag.Height)
                .DrawRectangle(Pens.DarkGray, 1, 2, Flag.Width - 2, Flag.Height - 2)
                Dim Format As StringFormat = New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
                FlagGraphics.DrawString(Text, TextFont, Brushes.Black, New Rectangle(0, 0, Flag.Width, Flag.Height), Format)
            End With
        End Using
        Return Flag

    End Function
#Region " RANDOM NUMBERS "
    Private ReadOnly Rnd As New Random()
    Public Function RandomNumber(ByVal Low As Integer, ByVal High As Integer) As Integer
        Dim EnsureLow = {Low, High}.Min
        Dim EnsureHigh = {Low, High}.Max
        Dim Random_Nbr As Integer = Rnd.Next(EnsureLow, EnsureHigh + 1)
        Return Random_Nbr
    End Function
    Public Function RandomBoolean() As Boolean
        Dim Value As Integer = RandomNumber(1, 100) Mod 2
        Return Value = 1
    End Function
    Public Function Shuffle(Items As IEnumerable(Of Object), Optional TakeCount As Integer = 0) As List(Of Object)

        Dim List As New List(Of Object)(Items)
        If Items Is Nothing Then
            Return Nothing

        ElseIf Items.Any Then
            For i = 1 To 100
                Dim RandomIndex As Integer = RandomNumber(0, List.Count - 1)
                Dim ListIndex As Integer = i Mod List.Count
                Dim ListItem = List(ListIndex)
                List.RemoveAt(ListIndex)
                List.Insert(RandomIndex, ListItem)
            Next
            If TakeCount = 0 Then
                Return List

            ElseIf TakeCount < 0 Then
                'Flag to use Random Count
                Return List.Take(RandomNumber(0, Items.Count - 1)).ToList

            Else
                Return List.Take(TakeCount).ToList

            End If

        Else
            Return List

        End If

    End Function
#End Region
    Public Enum RelativeCursor
        None
        Inside
        LeftOf
        RightOf
        Above
        Below
    End Enum
    Public Function CursorToControlPosition(ControlItem As Control, Optional RelativeBounds As Rectangle = Nothing) As RelativeCursor

        If ControlItem Is Nothing Then
            Return RelativeCursor.None

        ElseIf ControlItem.Visible Then
            Dim CursorPosition As Point = Cursor.Position
            Dim RelativePoint As Point = ControlItem.PointToScreen(RelativeBounds.Location)
            Dim RelativeSize As Size = If(RelativeBounds.Width = 0, ControlItem.Size, RelativeBounds.Size)
            Dim RelativeRectangle As New Rectangle(RelativePoint, RelativeSize)
            If RelativeRectangle.Contains(CursorPosition) Then
                Return RelativeCursor.Inside
            Else
                If CursorPosition.Y <= RelativeRectangle.Top Then
                    Return RelativeCursor.Above

                ElseIf CursorPosition.Y >= RelativeRectangle.Bottom Then
                    Return RelativeCursor.Below

                ElseIf CursorPosition.X <= RelativeRectangle.Left Then
                    Return RelativeCursor.LeftOf

                ElseIf CursorPosition.X >= RelativeRectangle.Right Then
                    Return RelativeCursor.RightOf

                Else
                    Return RelativeCursor.None

                End If
            End If
        Else
            Return RelativeCursor.None
        End If

    End Function
    Public Function CursorOverControl(ControlItem As Control) As Boolean
        Return CursorToControlPosition(ControlItem) = RelativeCursor.Inside
    End Function
    Public Function DateTimeToString(DateValue As Date) As String
        Return Format(DateValue, "M/d/yyyy HH:mm:ss.fff")
    End Function
    Public Function StringToDateTime(DateString As String) As Date

        If DateString Is Nothing Then
            Return New Date
        Else
            Dim DateValue As Date
            If Date.TryParseExact(DateString, "M/d/yyyy HH:mm:ss.fff", Nothing, DateTimeStyles.None, DateValue) Then
                Return DateValue
            Else
                Return New Date
            End If
        End If

    End Function
    Public Function DateToDB2Date(DateValue As Date) As String

        Dim Elements As New List(Of String) From {"'" + Format(DateValue.Year, "0000"),
            "-" + Format(DateValue.Month, "00"),
            "-" + Format(DateValue.Day, "00"),
            "'"}
        Return Join(Elements.ToArray, String.Empty)

    End Function
    Public Function DateToDB2Timestamp(DateValue As Date) As String

        Dim Elements As New List(Of String) From {"'" + Format(DateValue.Year, "0000"),
            "-" + Format(DateValue.Month, "00"),
            "-" + Format(DateValue.Day, "00"),
            "-" + Format(DateValue.Hour, "00"),
            "." + Format(DateValue.Minute, "00"),
            "." + Format(DateValue.Second, "00"),
            "." + Format(DateValue.Millisecond, "000000"),
            "'"}
        Return (Join(Elements.ToArray, String.Empty))

    End Function
    Public Function DB2TimestampToDate(Timestamp As String) As Date
        '2019-01-14-14.31.45.000304'
        Dim Values As New List(Of Integer)(From v In Regex.Matches(Timestamp, "[0-9]{2,}", RegexOptions.IgnoreCase) Select Integer.Parse(DirectCast(v, Match).Value, InvariantCulture))
        Return New DateTime(Values(0), Values(1), Values(2), Values(3), Values(4), Values(5), Values(6), DateTimeKind.Local)
    End Function
    Public Function TimespanToString(ElapsedValue As TimeSpan) As String

        Dim Elements As New List(Of String) From {"'" + Format(ElapsedValue.Hours, "00"),
            ":" + Format(ElapsedValue.Minutes, "00"),
            ":" + Format(ElapsedValue.Seconds, "00"),
            "." + Format(ElapsedValue.Milliseconds, "0000000"),
            "'"}
        Return (Join(Elements.ToArray, String.Empty))

    End Function
    Public Function DateToAccessString(DateValue As Date) As String

        '#4/1/2012#
        Dim Elements As New List(Of String) From {"#" + DateValue.Month.ToString(InvariantCulture),
            "/" + DateValue.Day.ToString(InvariantCulture),
            "/" + Format(DateValue.Year, "0000") + "#"}
        Return (Join(Elements.ToArray, String.Empty))

    End Function
    Public Function DB2ColumnNamingConvention(ColumnName As String) As String

        'https://social.msdn.microsoft.com/Forums/sqlserver/en-US/154c19c4-95ba-4b6f-b6ca-479288feabfb/characters-that-are-not-allowed-in-table-name-amp-column-name-in-sql-server-
        'Data Rule output column names are limited to 30 characters
        If If(ColumnName, String.Empty).Length = 0 Then
            Return String.Empty

        Else
            Dim Letters = ColumnName.ToArray.Take(30)
            Dim LetterIndex As Integer = 0
            Dim NewLetters As New List(Of Char)
            For Each Letter As Char In Letters
                If LetterIndex = 0 Then
                    'FIRST CAN BE _, @, #, A - Z
                    If Regex.Match(Letter, "[^A-Z@$#_]", RegexOptions.IgnoreCase).Success Then
                        NewLetters.Add(Chr(Asc("_")))
                    Else
                        NewLetters.Add(Letter)
                    End If

                Else
                    'SUBSEQUENT CAN BE _, @, #, A - Z, 0 - 9
                    If Regex.Match(Letter, "[^A-Z@$#_0-9]", RegexOptions.IgnoreCase).Success Then
                        NewLetters.Add(Chr(Asc("_")))
                    Else
                        NewLetters.Add(Letter)
                    End If

                End If
                LetterIndex += 1
            Next
            Return NewLetters.ToArray
        End If

    End Function
    Public Function GetWord(Line As String, Index As Integer, Optional LookForward As Boolean = True) As KeyValuePair(Of Integer, String)

        If Line Is Nothing Then
            Return Nothing
        Else
            Index = {0, {Index, Line.Length - 1}.Min}.Max
            If Line.Length = 0 Then
                Return Nothing

            Else
                Dim Words = RegexMatches(Line, "[^ ]{1,}", RegexOptions.IgnoreCase)
                If Words.Any Then
                    Dim Word As IEnumerable(Of Match)
                    Dim Letter As String = Line.Substring(Index, 1)
                    If Letter = " " Then
                        If LookForward Then
                            Word = From w In Words Where w.Index >= Index
                        Else
                            Word = From w In Words Order By w.Index Descending Where w.Index <= Index
                        End If
                    Else
                        Word = From w In Words Order By w.Index Descending Where w.Index <= Index
                    End If
                    If Word.Any Then
                        Return New KeyValuePair(Of Integer, String)(Word.First.Index, Word.First.Value)
                    Else
                        Return Nothing
                    End If
                Else
                    Return Nothing
                End If
            End If

        End If

    End Function
    Public Function RegexMatches(InputString As String, Pattern As String, Options As RegexOptions) As List(Of Match)

        If InputString Is Nothing Or Pattern Is Nothing Then
            Return Nothing
        Else
            Return (From m In Regex.Matches(InputString, Pattern, Options) Select DirectCast(m, Match)).ToList

        End If

    End Function
    Public Function RegexSplit(InputString As String, Pattern As String, Options As RegexOptions) As List(Of String)

        Dim UniqueDelimiter As String = "ﯔ"
        InputString = Regex.Replace(InputString, Pattern, UniqueDelimiter, Options)
        Return Split(InputString, UniqueDelimiter).ToList

    End Function
    Public Function RegexAlphaNumeric(InputString As String, Length As Integer, Optional IgnoreCase As RegexOptions = RegexOptions.None) As List(Of Match)

        Dim Matches = RegexMatches(InputString, AlphaNumericPattern, IgnoreCase)
        Return Matches.Where(Function(m) m.Value.Length = Length).ToList

    End Function
    Public Sub ToClipboard(Array As List(Of String))

        If Array IsNot Nothing AndAlso Array.Any Then
            Clipboard.SetText(Join(Array.ToArray, vbNewLine))
        Else
            Clipboard.SetText("ToClipboard=Empty array".ToString(InvariantCulture))
        End If

    End Sub
    Public Sub ToClipboard(Array As String())

        If Array IsNot Nothing AndAlso Array.Any Then
            Clipboard.SetText(Join(Array, vbNewLine))
        Else
            Clipboard.SetText("ToClipboard=Empty array".ToString(InvariantCulture))
        End If

    End Sub
    Public Function TrimReturn(InputString As String) As String
        Return Trim(Regex.Replace(InputString, "[\n\r]", String.Empty, RegexOptions.None))
    End Function
    Public Function WrapWords(InText As String, InFont As Font, Width As Integer) As Dictionary(Of Integer, String)

        If Not If(InText, String.Empty).Any Or InFont Is Nothing Or Width <= 0 Then
            Return Nothing

        Else
            Dim Words As New List(Of String)(From rm In RegexMatches(InText, "\s{0,1}[^\s]{1,}", RegexOptions.None) Select rm.Value)
            Dim LineBuilder As New Dictionary(Of Integer, String) From {
                {0, String.Empty}
                }
            For Each Word In Words
                Dim Line As String = LineBuilder.Last.Value
                Dim LineLength As Integer = MeasureText(Line, InFont).Width
                Dim WordSize As Size = TextRenderer.MeasureText(Word, InFont)

                Dim PrecededByCrb As Boolean = Regex.Match(Word, "(?<=[\n\r])[^\s]{1,}", RegexOptions.None).Success
                Dim TooLong As Boolean = LineLength + WordSize.Width > Width

                If PrecededByCrb Or TooLong Then
                    LineBuilder.Add(LineBuilder.Count, Regex.Match(Word, "[^\s]{1,}", RegexOptions.None).Value)      'Start of a new line...remove preceding space
                Else
                    LineBuilder(LineBuilder.Last.Key) &= Word
                End If
            Next
            Return LineBuilder

        End If

    End Function
    Public Function MeasureText(Text As String, TextFont As Font) As Size

        If Not If(Text, String.Empty).Any Or TextFont Is Nothing Then
            Return New Size(0, 0)

        Else
            Dim gTextSize As SizeF
            Using g As Graphics = Graphics.FromImage(My.Resources.Plus)
                Dim sf As New StringFormat With {.Trimming = StringTrimming.None}
                gTextSize = g.MeasureString(Text, TextFont, RectangleF.Empty.Size, sf)
            End Using
            Return New Size(CInt(gTextSize.Width), CInt(gTextSize.Height))

        End If

    End Function
    Public Function WrapText(Paragraphs As List(Of String), MaxSentenceLength As Integer) As List(Of String)

        If Paragraphs Is Nothing Then
            Return Nothing
        Else
            Dim Rows As New List(Of String)
            For Each Paragraph In Paragraphs
                Rows.Add(Join(WrapText(Paragraph, MaxSentenceLength).ToArray, Delimiter))
            Next
            Return Rows
        End If

    End Function
    Public Function WrapText(Paragraph As String, MaxSentenceLength As Integer) As List(Of String)

        Dim Lines As New List(Of String)
        Dim Line As New Dictionary(Of Integer, String)
        Dim Words = RegexMatches(Paragraph, "[^ §]{1,}", RegexOptions.None).OrderBy(Function(m) m.Index)
        For Each Word In Words
            If Lines.Any Then
                MaxSentenceLength = 70
            Else
                MaxSentenceLength = 40
            End If
            Dim WordIndex As Integer = Line.Count
            Line.Add(WordIndex, Word.Value)
            Dim WordString = Join(Line.Values.ToArray, " ")
            If WordString.Length > MaxSentenceLength Then
                Line.Remove(WordIndex)
                Lines.Add(Join(Line.Values.ToArray, " "))
                Line.Clear()
                Line.Add(Line.Count, Word.Value)

            ElseIf Word Is Words.Last Then
                Lines.Add(Join(Line.Values.ToArray, " "))
                Line.Clear()

            End If
        Next
        Return Lines

    End Function
    Public Function ProperCase(Line As String) As String

        If Line Is Nothing Then
            Return Nothing
        Else
            Dim culture_info As CultureInfo = Threading.Thread.CurrentThread.CurrentCulture
            Dim text_info As TextInfo = culture_info.TextInfo
            Return text_info.ToTitleCase(Line.ToLower(culture_info))
        End If

    End Function
    Public Function CenterItem(ItemSize As Size) As Point

        Dim ScreenCenter As New Point(Convert.ToInt32(WorkingArea.Width / 2), Convert.ToInt32(WorkingArea.Height / 2))
        Dim ObjectCenter As New Size(Convert.ToInt32(ItemSize.Width / 2), Convert.ToInt32(ItemSize.Height / 2))
        ScreenCenter.Offset(-ObjectCenter.Width, -ObjectCenter.Height)
        Return ScreenCenter

    End Function
    Public Function GetHexColor(ColorName As Color) As String
        Return "#" & Hex(ColorName.R) & Hex(ColorName.G) & Hex(ColorName.B)
    End Function
    Public Function Bulletize(Items As String()) As String

        Dim List As New List(Of String)
        If Items IsNot Nothing Then
            For Each Item In Items
                List.Add("● " & Item)
            Next
        End If
        Return Join(List.ToArray, vbNewLine)

    End Function
    Public Function Bulletize(Items As Dictionary(Of String, List(Of String))) As String

        Dim List As New List(Of String)
        If Items IsNot Nothing Then
            For Each Item In Items.Keys
                List.Add("± " & Item)
                For Each SubItem In Items(Item)
                    List.Add("    ● " & SubItem)
                Next
            Next
        End If
        Return Join(List.ToArray, vbNewLine)

    End Function
    Public Function Guestimate() As Integer

        '|          |     X              | X=Width/2...wrapped=False
        '|        Y |     | Y=X/2...wrapped=true
        '|          |  Z  | Z=Y

        Dim Answer As Integer = 129
        Dim Lefts As New List(Of Integer)({0})
        Dim Rights As New List(Of Integer)({497})
        Dim Attempts As New List(Of Integer)

        Do
            Dim Delta As Integer = Rights.Min - Lefts.Max
            Dim Mid As Integer = Lefts.Max + Convert.ToInt32(Delta / 2)
            'If Attempts.Count = 2 Then Stop
            Attempts.Add(Mid)
            If Rights.Min - Mid <= 1 Then
                Exit Do
            Else
                If Mid < Answer Then
                    Lefts.Add(Mid)

                Else
                    Rights.Add(Mid)

                End If
            End If
            If Attempts.Count > Rights.Max Then Exit Do

        Loop
        Return Attempts.Last

    End Function
    Public Function LevenshteinDistance(ByVal s As String, ByVal t As String) As Integer

        If s Is Nothing Or t Is Nothing Then
            Return 0
        Else
            Dim n As Integer = s.Length
            Dim m As Integer = t.Length
            Dim d As New Dictionary(Of Point, Integer)      '(n + 1, m + 1) As Integer

            If n = 0 Then
                Return m
            Else
                If m = 0 Then
                    Return n
                Else
                    For i As Integer = 0 To n
                        d.Add(New Point(i, 0), i)
                    Next
                    For j As Integer = 0 To m
                        d.Add(New Point(0, j), j)
                    Next
                    For i As Integer = 1 To n
                        For j As Integer = 1 To m
                            Dim Cost As Integer
                            If t(j - 1) = s(i - 1) Then
                                Cost = 0
                            Else
                                Cost = 1
                            End If
                            d(New Point(i, j)) = Math.Min(Math.Min(d(New Point(i - 1, j)) + 1, d(New Point(i, j - 1))) + 1, d(New Point(i - 1, j - 1)) + Cost)
                        Next
                    Next
                    Return d(New Point(n, m))
                End If
            End If
        End If

    End Function
    Friend Sub ParenthesisNodes(StringNode As StringData, TextIn As String)

        REM /// MUST BE DELIMITED BY A CHARACTER WHICH WILL NEVER BE FOUND IN SCRIPT
        Dim Group = ParenthesisCapture(TextIn)

        REM /// GROUP.LENGTH=0 MEANS NO () FOUND IN TextIn
        REM /// IF TextIn.LENGTH=0 THEN HAS REACHED EOL

        If TextIn.Length = 0 Then
            REM /// EOL

        ElseIf Group.Length = 0 Then
            REM /// NO PARENTHESIS LEFT=SIBLINGS ADDED, NOW ADD CHILDREN
            Dim NewNodes As List(Of StringData) = StringNode.Parentheses
            For Each ChildNode In NewNodes
                Dim TextValues = Split(ChildNode.Value, NonCharacter)
                ChildNode.Value = TextValues.First
                ParenthesisNodes(ChildNode, TextValues.Last)
            Next

        Else
            REM /// FOUND PARENTHESIS...ADD SIBLINGS BY RECURSING ON TEXT. MUST SUBSTITUTE PARENTHESIS WITH {} OTHERWISE INFINITE LOOP
            Dim ChildText As String = "{" & TextIn.Substring(Group.Start + 1, Group.Length - 2) & "}"
            Dim SiblingText As String = TextIn.Remove(Group.Start, Group.Length)
            SiblingText = SiblingText.Insert(Group.Start, StrDup(Group.Length, "-"))

            Dim NodeText As String = Join({Group.Value, NonCharacter, ChildText}, String.Empty)
            If StringNode IsNot Nothing Then
                Dim ParentGroup As StringData = StringNode
                Group.Start += ParentGroup.Start
                StringNode.Parentheses.Add(New StringData With {
                                      .Start = Group.Start,
                                      .Length = Group.Length,
                                      .Value = NodeText})
            End If
            ParenthesisNodes(StringNode, SiblingText)

        End If

    End Sub
    Public Function ParenthesisCapture(Text As String) As StringData

        Dim Capture As New StringData
        With Capture
            If Text Is Nothing Then
            Else
                Dim Parentheses As New List(Of Match)(From M In Regex.Matches(Text, "\(|\)", RegexOptions.IgnoreCase) Select DirectCast(M, Match))
                Dim LeftCount As Integer = 0, RightCount As Integer = 0
                For Each Parenthese In Parentheses
                    If Parenthese.Value = "(" Then LeftCount += 1
                    If Parenthese.Value = ")" Then RightCount += 1
                    If LeftCount = RightCount Then
                        .Start = Parentheses.First.Index
                        .Length = 1 + Math.Abs(Parenthese.Index - Parentheses.First.Index)
                        .Value = Text.Substring(.Start, .Length)
                        Exit For
                    End If
                Next
            End If
        End With
        Return Capture

    End Function
#Region " ENUMS "
    Public Function EnumNames(EnumType As Type) As List(Of String)
        Return [Enum].GetNames(EnumType).ToList
    End Function
    Public Function ParseEnum(Of T)(ByVal value As String) As T
        Dim Names As New List(Of String)(EnumNames(GetType(T)))
        If Names.Contains(value) Then
            Return CType([Enum].Parse(GetType(T), value, True), T)
        Else
            Return Nothing
        End If
    End Function
#End Region
    Public Function ColorImages() As Dictionary(Of String, Image)

        Dim ColorImageCollection As New Dictionary(Of String, Image)
        REM /// INITIALIZE THEM
        Dim X As Color = Color.Beige
        Dim ColorType As Type = X.GetType
        Dim ColorList() As PropertyInfo = ColorType.GetProperties(BindingFlags.Static Or BindingFlags.DeclaredOnly Or BindingFlags.Public)
        Dim Colors As New List(Of String)(From CL In ColorList Select CL.Name)
        Dim ComboImage As Image = Nothing
        For Each ItemColor In Colors
            Dim _Image As New Bitmap(16, 16)
            Using G As Drawing.Graphics = Drawing.Graphics.FromImage(_Image)
                Using Brush As New SolidBrush(Color.FromName(ItemColor))
                    G.DrawRectangle(Pens.Black, 0, 0, _Image.Width - 1, _Image.Height - 1)
                    G.FillRectangle(Brush, 2, 2, _Image.Width - 4, _Image.Height - 4)
                End Using
            End Using
            ColorImageCollection.Add(ItemColor, _Image)
        Next
        Return ColorImageCollection

    End Function
    Public Function ChangeImageColor(ByVal bmp As Bitmap, ByVal OldColor As Color, NewColor As Color) As Image

        If bmp IsNot Nothing Then
            Using g As Graphics = Graphics.FromImage(bmp)
                g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                Dim ColorMap As ColorMap() = New ColorMap(0) {}
                ColorMap(0) = New ColorMap With {
                    .OldColor = OldColor,
                    .NewColor = NewColor
                }
                Using Attributes As ImageAttributes = New ImageAttributes()
                    Attributes.SetRemapTable(ColorMap)
                    Dim rect As Rectangle = New Rectangle(0, 0, bmp.Width, bmp.Height)
                    g.DrawImage(bmp, rect, 0, 0, rect.Width, rect.Height, GraphicsUnit.Pixel, Attributes)
                End Using
            End Using
        End If
        Return bmp

    End Function
    Public Function DrawRoundedRectangle(ByVal Rect As Rectangle, Optional ByVal Corner As Integer = 10) As System.Drawing.Drawing2D.GraphicsPath

        Dim Graphix As New System.Drawing.Drawing2D.GraphicsPath
        Dim ArcRect As New RectangleF(Rect.Location, New SizeF(Corner, Corner))
        Graphix.AddArc(ArcRect, 180, 90)
        Graphix.AddLine(Rect.X + CInt(Corner / 2), Rect.Y, Rect.X + Rect.Width - CInt(Corner / 2), Rect.Y)
        ArcRect.X = Rect.Right - Corner
        Graphix.AddArc(ArcRect, 270, 90)
        Graphix.AddLine(Rect.X + Rect.Width, Rect.Y + CInt(Corner / 2), Rect.X + Rect.Width, Rect.Y + Rect.Height - CInt(Corner / 2))
        ArcRect.Y = Rect.Bottom - Corner
        Graphix.AddArc(ArcRect, 0, 90)
        Graphix.AddLine(Rect.X + CInt(Corner / 2), Rect.Y + Rect.Height, Rect.X + Rect.Width - CInt(Corner / 2), Rect.Y + Rect.Height)
        ArcRect.X = Rect.Left
        Graphix.AddArc(ArcRect, 90, 90)
        Graphix.AddLine(Rect.X, Rect.Y + CInt(Corner / 2), Rect.X, Rect.Y + Rect.Height - CInt(Corner / 2))
        Return Graphix

    End Function
    Friend Function SetOpacity(ByVal image As Image, ByVal opacity As Single) As Image

        Dim output = New Bitmap(image.Width, image.Height)
        Dim colorMatrix = New ColorMatrix With {
            .Matrix33 = opacity
        }
        Using imageAttributes As New ImageAttributes
            imageAttributes.SetColorMatrix(colorMatrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap)
            Using gfx = Graphics.FromImage(output)
                gfx.SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias
                gfx.DrawImage(image, New Rectangle(0, 0, image.Width, image.Height), 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, imageAttributes)
            End Using
        End Using
        Return output

    End Function
    Public Function KeyIsDown(Key As Keys) As Boolean
        '-32767=Down
        '0 Up
        '1 Up+Released
        Return Not {0, 1}.Contains(NativeMethods.GetKeyState(Key))
    End Function
#Region " FILES "
    Public Sub ZipIt(ZipPath As String, ZipFiles As String())

        If ZipFiles IsNot Nothing Then
            If File.Exists(ZipPath) Then
                File.Delete(ZipPath)
            End If
            Using Zip As ZipArchive = ZipFile.Open(ZipPath, ZipArchiveMode.Create)
                For Each File In ZipFiles
                    Zip.CreateEntryFromFile(File, Path.GetFileName(File), CompressionLevel.Optimal)
                Next
            End Using
        End If

    End Sub
    Public Function WriteText(FilePathOrName As String, List As List(Of String)) As Boolean

        If List Is Nothing Then
            Return False
        Else
            Dim Content As String = Join(List.ToArray, vbNewLine)
            Return WriteText(FilePathOrName, Content)
        End If

    End Function
    Public Function WriteText(FilePathOrName As String, Items As String()) As Boolean

        If Items Is Nothing Then
            Return False
        Else
            Dim Content As String = Join(Items, vbNewLine)
            Return WriteText(FilePathOrName, Content)
        End If

    End Function
    Public Function WriteText(FilePathOrName As String, Content As String) As Boolean

        If IsFile(FilePathOrName) Then
            If File.Exists(FilePathOrName) Then
                Using SW As New StreamWriter(FilePathOrName)
                    SW.Write(Content)
                End Using
                Return True
            Else
                Return False
            End If
        Else
            Dim TryDesktop As String = Desktop & "\" & FilePathOrName & ".txt"
            If File.Exists(TryDesktop) Then
                Using SW As New StreamWriter(TryDesktop)
                    SW.Write(Content)
                End Using
                Return True
            Else
                Return False
            End If
        End If

    End Function
    Public Function ReadText(FilePathOrName As String) As String

        Dim CanRead As Boolean = IsFile(FilePathOrName) And File.Exists(FilePathOrName)
        If Not CanRead Then                 'Try cleaning up provided value
            'Could be Name only as ABC
            'Could be Name + Extension as ABC.txt
            'Could be Fullpath, but no extension as C:\Users\SEANGlover\Desktop\PSRR\DDL_SQL\ABC
            Dim kvp = GetFileNameExtension(FilePathOrName)
            If IsFile(FilePathOrName) Then
                If kvp.Value = Extensions.None Then
                    FilePathOrName &= ".txt"
                Else
                    'Is a filepath and has extension ( valid or not ) however does not exist at location
                    Return Nothing
                End If

            Else
                'Not a file format so could be Name only as ABC Or Name + Extension as ABC.txt...Assume to Desktop
                FilePathOrName = Desktop & "\" & FilePathOrName
                If kvp.Value = Extensions.None Then FilePathOrName &= ".txt"

            End If
            'Try again after cleanup
            CanRead = IsFile(FilePathOrName) And File.Exists(FilePathOrName)
        End If
        If CanRead Then
            Dim Content As String
            Using SR As New StreamReader(FilePathOrName)
                Content = SR.ReadToEnd
            End Using
            Return Content
        Else
            Return Nothing
        End If

    End Function
    Public Function IsFile(Source As String) As Boolean
        Return Regex.Match(Source, FilePattern, RegexOptions.IgnoreCase).Success
    End Function
    <Flags()> Public Enum Extensions
        None
        Invalid
        Excel
        Text
        CommaSeparated
        PortableDocumentFormat
        SQL
        Unknown
    End Enum
    Public Function MyTextFiles(Optional Manager As Resources.ResourceManager = Nothing) As Dictionary(Of String, String)

        Manager = If(Manager, My.Resources.ResourceManager)
        Dim ResourceSet As Resources.ResourceSet = Manager.GetResourceSet(CultureInfo.CurrentCulture, True, True)
        Dim Resources As New List(Of DictionaryEntry)(ResourceSet.OfType(Of DictionaryEntry).Where(Function(x) x.Value.GetType Is GetType(String)))
        Return Resources.ToDictionary(Function(x) x.Key.ToString, Function(y) y.Value.ToString)

    End Function
    Public Function MyImages(Optional Manager As Resources.ResourceManager = Nothing) As Dictionary(Of String, Image)

        Manager = If(Manager, My.Resources.ResourceManager)
        Dim ResourceSet As Resources.ResourceSet = Manager.GetResourceSet(CultureInfo.CurrentCulture, True, True)
        Dim Resources As New List(Of DictionaryEntry)(ResourceSet.OfType(Of DictionaryEntry).Where(Function(x) x.Value.GetType Is GetType(Bitmap)))
        Return Resources.ToDictionary(Function(x) x.Key.ToString, Function(y) DirectCast(y.Value, Image))

    End Function
    Public Function MyIcons(Optional Manager As Resources.ResourceManager = Nothing) As Dictionary(Of String, Icon)

        Manager = If(Manager, My.Resources.ResourceManager)
        Dim ResourceSet As Resources.ResourceSet = Manager.GetResourceSet(CultureInfo.CurrentCulture, True, True)
        Dim Resources As New List(Of DictionaryEntry)(ResourceSet.OfType(Of DictionaryEntry).Where(Function(x) x.Value.GetType Is GetType(Icon)))
        Return Resources.ToDictionary(Function(x) x.Key.ToString, Function(y) DirectCast(y.Value, Icon))

    End Function
    Public Function GetFiles(Path As String, Extension As String) As List(Of String)

        Return (From Folder In SafeWalk.EnumerateFiles(Path, "*" & Extension, SearchOption.AllDirectories)).ToList

    End Function
    Public Function GetFiles(Path As String, Extension As String, Level As SearchOption) As List(Of String)

        Return (From Folder In SafeWalk.EnumerateFiles(Path, "*" & Extension, Level)).ToList

    End Function
    Public Function GetFileNameExtension(Path As String) As KeyValuePair(Of String, Extensions)

        Dim NameAndFilter As String = Split(Path, "\").Last
        Dim FileNameExtension As String() = Split(NameAndFilter, ".")
        Dim FileName As String = FileNameExtension.First

        If FileNameExtension.Count = 1 Then
            'Missing extension
            Return New KeyValuePair(Of String, Extensions)(FileName, Extensions.None)

        ElseIf FileNameExtension.Count = 2 Then
            Dim FileFilter As String = FileNameExtension.Last.ToUpperInvariant
            Select Case True
                Case FileFilter.StartsWith("XL", StringComparison.InvariantCulture)
                    Return New KeyValuePair(Of String, Extensions)(FileName, Extensions.Excel)

                Case FileFilter = "TXT"
                    Return New KeyValuePair(Of String, Extensions)(FileName, Extensions.Text)

                Case FileFilter = "CSV"
                    Return New KeyValuePair(Of String, Extensions)(FileName, Extensions.CommaSeparated)

                Case FileFilter = "SQL"
                    Return New KeyValuePair(Of String, Extensions)(FileName, Extensions.SQL)

                Case FileFilter = "PDF"
                    Return New KeyValuePair(Of String, Extensions)(FileName, Extensions.PortableDocumentFormat)

                Case Else
                    Return New KeyValuePair(Of String, Extensions)(FileName, Extensions.Unknown)

            End Select

        Else
            'Can never have 2 or more . in a filepath
            Return New KeyValuePair(Of String, Extensions)(FileName, Extensions.Invalid)
        End If

    End Function
    Public Function PathToList(FilePath As String) As List(Of String)

        If FilePath Is Nothing Then
            Return Nothing
        Else
            Dim Content As String = ReadText(FilePath)
            Dim Lines As New List(Of String)(From line In Split(Content, vbNewLine) Where Trim(line).Any)
            Return Lines
        End If

    End Function
    Public Function ExcelSheetNames(Location As String) As List(Of String)

        Dim Sheets As New List(Of String)
        Dim ExcelConnectionACE As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Location & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1;"""
        Dim ExcelConnectionJet As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Location & ";Extended Properties=""Excel 8.0;HDR=Yes;"""

        Dim Filter As String = Split(Location, ".").Last
        Dim ExcelConnectionString As String = If(Filter = "xls", ExcelConnectionJet, ExcelConnectionACE)

        Using ExcelConnection As New OleDbConnection(ExcelConnectionString)
            Try
                ExcelConnection.Open()
                Dim ExcelTables As DataTable = ExcelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing})
                Sheets = (From ET In ExcelTables Where Not ET("TABLE_NAME").ToString.ToUpperInvariant.Contains("FILTERDATABASE") Select ET("TABLE_NAME").ToString.ToUpperInvariant).ToList

            Catch FileInUseException As OleDbException
                Using Message As New Prompt
                    Message.Show("File in use", FileInUseException.Message)
                End Using
            End Try

        End Using
        Return Sheets

    End Function
    Public Function DataColumnToList(Column As DataColumn) As List(Of Object)

        Dim Objects As New List(Of Object)
        If Column IsNot Nothing Then Objects = (From r In Column.Table.AsEnumerable Select r(Column)).ToList
        Return Objects

    End Function
#End Region
#Region " ENCRYPTION "
    Public Function Krypt(TextIn As String) As String
        Return Convert.ToBase64String(System.Text.Encoding.Unicode.GetBytes(TextIn))
    End Function
    Public Function DeKrypt(TextIn As String) As String
        Return System.Text.Encoding.Unicode.GetString(Convert.FromBase64String(TextIn))
    End Function
#End Region
#Region " VALUE TYPES "
    Public Function GetDataType(Column As DataColumn) As Type
        Return GetDataType(DataColumnToList(Column))
    End Function
    Public Function GetDataType(Types As List(Of Type)) As Type

        If Types Is Nothing Then
            Return Nothing
        Else
#Region " STRING AS DEFAULT "
            If Not Types.Any Then
                Return GetType(String)
#End Region
#Region " ONE INSTANCE OF STRING MUST RETURN STRING "
            ElseIf Types.Intersect({GetType(String)}).Count = 1 Then
                Return GetType(String)
#End Region
#Region " ALL ARE OF DATE "
            ElseIf Types.Intersect({GetType(Date), GetType(DateAndTime)}).Count = Types.Count Then
                Return GetType(Date)
#End Region
#Region " ALL ARE OF BOOLEAN "
            ElseIf Types.Intersect({GetType(Boolean)}).Count = Types.Count Then
                Return GetType(Boolean)
#End Region
#Region " EACH TYPE IS A WHOLE NUMBER "
            ElseIf Types.Intersect({GetType(Byte), GetType(Short), GetType(Integer), GetType(Long)}).Count = Types.Count Then
                REM /// DESCENDING IN SIZE
                If Types.Contains(GetType(Long)) Then
                    Return GetType(Long)

                ElseIf Types.Contains(GetType(Integer)) Then
                    Return GetType(Integer)

                ElseIf Types.Contains(GetType(Short)) Then
                    Return GetType(Short)

                Else
                    Return GetType(Byte)
                End If
#End Region
#Region " EACH TYPE IS NUMERIC AND NOT ALL ARE WHOLE "
            ElseIf Types.Intersect({GetType(Byte), GetType(Short), GetType(Integer), GetType(Long), GetType(Double), GetType(Decimal)}).Count = Types.Count Then
                Return GetType(Double)
#End Region
#Region " EACH TYPE IS EITHER BITMAP Or IMAGE "
            ElseIf Types.Intersect({GetType(Bitmap), GetType(Image)}).Count = Types.Count Then
                Return GetType(Image)
#End Region
#Region " EACH TYPE IS AN ICON "
            ElseIf Types.Intersect({GetType(ICON)}).Count = Types.Count Then
                Return GetType(Icon)
#End Region
#Region " MIXED TYPES "
            Else
                Return GetType(String)
#End Region
            End If
        End If

    End Function
    Public Function GetDataType(Values As List(Of Object)) As Type

        If Values Is Nothing Then
            Return Nothing
        Else
            Dim Types = From V In Values Where Not (IsDBNull(V) Or IsNothing(V)) Select GetDataType(V.ToString)
            Dim BlendedType = GetDataType(Types.Distinct)
            'If (From v In Values Where v.GetType Is GetType(Bitmap)).Any Then Stop
            'If (From t In Types Where t Is GetType(Image)).Any Then Stop
            Return BlendedType
        End If

    End Function
    Public Function GetDataType(Types As IEnumerable(Of Type)) As Type

        If Types Is Nothing Then
            Return Nothing
        Else
            'If (From t In Types Where t Is GetType(Image)).Any Then Stop
            Return GetDataType(Types.ToList)
        End If

    End Function
    Public Function GetDataType(Types As List(Of String)) As Type

        If Types Is Nothing Then
            Return Nothing
        Else
            Return GetDataType((From t In Types Select GetDataType(t)).Distinct.ToList)
        End If

    End Function
    Public Function GetDataType(Value As Object) As Type
        If Value Is Nothing Then
            Return Nothing
        Else
            Return GetDataType(Value.ToString)
        End If
    End Function
    Public Function GetDataType(Value As String) As Type

        If Value Is Nothing Then
            Return GetType(String)

        Else
            If Value.Contains("Drawing.Bitmap") Or Value.Contains("Drawing.Image") Then
                Return GetType(Image)

            ElseIf Value.Contains("Drawing.Icon") Then
                Return GetType(Icon)

            Else
                Dim _Date As Date
                Dim Formats() As String = {
                    "M/d/yyyy",
                    "M/d/yyyy h:mm",
                    "M/d/yyyy h:mm:ss",
                    "M/d/yyyy h:mm:ss tt"}

                If Date.TryParseExact(Value, Formats, New CultureInfo("en-US"), DateTimeStyles.AllowWhiteSpaces, _Date) Then
                    Return _Date.GetType

                Else
                    Dim _Boolean As Boolean
                    If Boolean.TryParse(Value, _Boolean) Or Value.ToUpperInvariant = "TRUE" Or Value.ToUpperInvariant = "FALSE" Then
                        Return _Boolean.GetType

                    Else
                        If IsNumeric(Value) Then
                            Dim _Decimal As Decimal
                            If Decimal.TryParse(Value, _Decimal) Then
                                REM /// NUMERIC+COULD BE DECIMAL Or INTEGER
                                If Split(Value, ".").Count = 1 Then
                                    REM /// INTEGER
                                    REM /// MUST BE A WHOLE NUMBER. START WITH SMALLEST AND WORK UP
                                    Dim _Byte As Byte
                                    If Byte.TryParse(Value, _Byte) Then
                                        Return _Byte.GetType

                                    Else
                                        Dim _Short As Short
                                        If Short.TryParse(Value, _Short) Then
                                            Return _Short.GetType

                                        Else
                                            Dim _Integer As Integer
                                            If Integer.TryParse(Value, _Integer) Then
                                                Return _Integer.GetType

                                            Else
                                                Dim _Long As Long
                                                If Long.TryParse(Value, _Long) Then
                                                    Return _Long.GetType

                                                Else
                                                    REM /// NOT DATE, BOOLEAN, DECIMAL, NOR INTEGER...DEFAULT TO STRING
                                                    Return GetType(String)

                                                End If
                                            End If
                                        End If
                                    End If

                                Else
                                    REM /// DECIMAL
                                    Return _Decimal.GetType

                                End If
                            Else
                                REM /// NUMERIC+COULD BE SCIENTIFIC *** NEED CODE TO CONVERT SCIENTIFIC STRING TO PROPER NUMBER, THEN FEED BACK INTO DECIMAL/INTEGER DETERMINATION
                                Return GetType(String)

                            End If
                        Else
                            REM /// NOT DATE, BOOLEAN, NOR NUMERIC...DEFAULT TO STRING
                            Return GetType(String)

                        End If
                    End If
                End If
            End If
        End If

    End Function
#End Region
    Public Enum HandlerAction
        Add
        Remove
    End Enum
    Public Function Desktoptxtpath(Name As String) As String
        Return Join({Desktop, "/", Name, ".txt"}, String.Empty)
    End Function
    Public Function MyDocumentstxtpath(Name As String) As String
        Return Join({MyDocuments, "/", Name, ".txt"}, String.Empty)
    End Function
    Public Sub Wait(Miliseconds As Long)

        Dim SW As New Stopwatch
        SW.Start()
        Do Until SW.ElapsedMilliseconds >= Miliseconds
        Loop
        SW.Stop()

    End Sub
End Module

Namespace Pdf2Text
    Public Class ConversionEventArgs
        Inherits EventArgs
        Public ReadOnly Property Succeeded As Boolean
        Public ReadOnly Property Message As String
        Public Sub New(Success As Boolean, Optional Message As String = Nothing)
            Succeeded = Success
            Me.Message = Message
        End Sub
    End Class
    Public Class ConversionCollection
        Inherits List(Of Conversion)
        Public Event Completed(sender As Object, e As ConversionEventArgs)
        Public Event ItemCompleted(sender As Object, e As ConversionEventArgs)
        Public ReadOnly Property Items As List(Of String)
        Public ReadOnly Property SourceFolder As String
        Public ReadOnly Property DestinationFolder As String
        Public ReadOnly Property Started As New Date
        Public ReadOnly Property Ended As New Date
        Public ReadOnly Property Succeeded As Boolean
            Get
                Return Where(Function(p) p.Succeeded).Count = Count
            End Get
        End Property
        Public Sub New(SourceItems As List(Of String), DestinationFolder As String)
            'Get only provided items
            _Items = SourceItems
            _DestinationFolder = DestinationFolder
            Fill()
        End Sub
        Public Sub New(Items As String())
            'Get only provided items + Destination MUST be provided in With {.DestinationFolder=""}
            If Items Is Nothing Then

            Else
                _Items = Items.ToList
                Fill()
            End If

        End Sub
        Public Sub New(Items As List(Of String))
            'Get only provided items + Destination MUST be provided in With {.DestinationFolder=""}
            _Items = Items
            Fill()
        End Sub
        Public Sub New(SourceFolder As String)
            _SourceFolder = SourceFolder
            _DestinationFolder = SourceFolder
            'All items in a Folder
            _Items = GetFiles(SourceFolder, ".pdf")
            Fill()
        End Sub
        Public Sub New(SourceFolder As String, DestinationFolder As String)
            _SourceFolder = SourceFolder
            _DestinationFolder = DestinationFolder
            'All items in a Folder
            _Items = GetFiles(SourceFolder, ".pdf")
            Fill()
        End Sub
        Private Sub Fill()

            For Each pdf In Items
                Dim pdf_Filename As String = Split(pdf, "\").Last
                Dim Destination_Path As String = DestinationFolder & Replace(pdf_Filename, ".pdf", ".txt")
                Add(New Conversion(pdf, Destination_Path))
                Last.Parent = Me
                AddHandler Last.Completed, AddressOf ConversionCompleted
            Next

        End Sub
        Public Sub StartConversions()

            _Started = Now
            If Count = 0 Then
                _Ended = _Started
                RaiseEvent Completed(Me, New ConversionEventArgs(False))
            Else
                For Each pdfItem In Me
                    pdfItem.Convert()
                Next
            End If

        End Sub
        Private Sub ConversionCompleted(sender As Object, e As ConversionEventArgs)

            With DirectCast(sender, Conversion)
                RemoveHandler .Completed, AddressOf ConversionCompleted
                RaiseEvent ItemCompleted(sender, New ConversionEventArgs(.Succeeded))
                If Where(Function(p) p.Ended > New Date).Count = Count Then
                    _Ended = Now
                    RaiseEvent Completed(Me, New ConversionEventArgs(.Succeeded))
                End If
            End With

        End Sub
    End Class
    Public Class Conversion
        Public Event Completed(sender As Object, e As ConversionEventArgs)
        Public ReadOnly Property PdfPath As String
        Public ReadOnly Property TxtPath As String
        Friend Property Parent As ConversionCollection
        Public ReadOnly Property Index As Integer
            Get
                If Parent Is Nothing Then
                    Return 0
                Else
                    Return Parent.IndexOf(Me)
                End If

            End Get
        End Property
        Public ReadOnly Property Started As New Date
        Public ReadOnly Property Ended As New Date
        Public ReadOnly Property Succeeded As Boolean
        Public ReadOnly Property Content As String
        Public Sub New(pdfPath As String)

            _PdfPath = pdfPath
            _TxtPath = Replace(pdfPath, ".pdf", ".txt")

        End Sub
        Public Sub New(pdfPath As String, txtPath As String)

            _PdfPath = pdfPath
            _TxtPath = txtPath

        End Sub
        Public Sub Convert()

            If File.Exists(PdfPath) Then
                With New BackgroundWorker
                    AddHandler .DoWork, AddressOf PdfWorker_DoWork
                    AddHandler .RunWorkerCompleted, AddressOf PdfWorker_Completed
                    Do While .IsBusy
                    Loop
                    .RunWorkerAsync()
                End With
            Else
                _Started = Now
                _Ended = _Started
                _Succeeded = False
                RaiseEvent Completed(Me, New ConversionEventArgs(False, PdfPath & " not found"))
            End If

        End Sub
        Private Sub PdfWorker_DoWork(sender As Object, e As EventArgs)

            _Started = Now
            With DirectCast(sender, BackgroundWorker)
                RemoveHandler .DoWork, AddressOf PdfWorker_DoWork

            End With
            Dim doc As PDDocument = Nothing
            Try
                doc = PDDocument.load(PdfPath)
                Dim Stripper As New PDFTextStripper()
                Dim txtString = Stripper.getText(doc)
                _Succeeded = True
                _Content = txtString
                Using sw As StreamWriter = New StreamWriter(TxtPath)
                    sw.WriteLine(Content)
                End Using

            Catch ex As Exception       ' java.io.IOException
                _Succeeded = False
                _Content = String.Empty

            Finally
                If doc IsNot Nothing Then
                    doc.close()
                End If

            End Try
            _Ended = Now

        End Sub
        Private Sub PdfWorker_Completed(sender As Object, e As RunWorkerCompletedEventArgs)

            With DirectCast(sender, BackgroundWorker)
                RemoveHandler .RunWorkerCompleted, AddressOf PdfWorker_Completed
            End With
            _Ended = Now
            RaiseEvent Completed(Me, New ConversionEventArgs(True))

        End Sub
    End Class
End Namespace

Namespace TLP
    Public Module Sizing
        Public Function GetBorderThickness(Border As BorderStyle) As Integer
            Dim BorderThicknesses As New Dictionary(Of BorderStyle, Integer)
            BorderThicknesses.Clear()
            BorderThicknesses.Add(BorderStyle.None, 0)
            BorderThicknesses.Add(BorderStyle.FixedSingle, 3)
            BorderThicknesses.Add(BorderStyle.Fixed3D, 5)
            Return BorderThicknesses(Border)
        End Function
        Public Function GetCellBorderThickness(Border As TableLayoutPanelCellBorderStyle) As Integer
            Dim CellBorderThicknesses As New Dictionary(Of TableLayoutPanelCellBorderStyle, Integer)
            CellBorderThicknesses.Clear()
            CellBorderThicknesses.Add(TableLayoutPanelCellBorderStyle.None, 0)
            CellBorderThicknesses.Add(TableLayoutPanelCellBorderStyle.Single, 1)
            CellBorderThicknesses.Add(TableLayoutPanelCellBorderStyle.Inset, 2)
            CellBorderThicknesses.Add(TableLayoutPanelCellBorderStyle.Outset, 2)
            CellBorderThicknesses.Add(TableLayoutPanelCellBorderStyle.InsetDouble, 3)
            CellBorderThicknesses.Add(TableLayoutPanelCellBorderStyle.OutsetDouble, 3)
            CellBorderThicknesses.Add(TableLayoutPanelCellBorderStyle.OutsetPartial, 3)
            Return CellBorderThicknesses(Border)
        End Function
        Public Function GetContentSpace(TLP As TableLayoutPanel) As Integer

            If TLP IsNot Nothing Then
                Dim Values As New List(Of Integer)
                With TLP
                    'Left BorderSide
                    Values.Add(GetBorderThickness(.BorderStyle))
                    For i = 0 To TLP.ColumnCount
                        Values.Add(GetCellBorderThickness(.CellBorderStyle))
                    Next
                    'Right BorderSide
                    Values.Add(GetBorderThickness(.BorderStyle))
                End With
                Return TLP.Width - Values.Sum
            Else
                Return 0
            End If

        End Function
        Public Function GetSize(TLP As TableLayoutPanel) As Size

            If TLP IsNot Nothing Then
                Dim BorderThickness As Integer = GetBorderThickness(TLP.BorderStyle)
                Dim CellBorderThickness As Integer = GetCellBorderThickness(TLP.CellBorderStyle)

                REM /// MAKE COLUMN STRIPS AND ROW STRIPS TO GET A MAX WIDTH / HEIGHT VALUE FOR AUTOSIZE COLUMNS AND ROWS
                Dim Columns As New Dictionary(Of Integer, List(Of Control))
                Dim Rows As New Dictionary(Of Integer, List(Of Control))
                For Each Item As Control In TLP.Controls
                    Dim xy As TableLayoutPanelCellPosition = TLP.GetCellPosition(Item)
                    If Not Columns.Keys.Contains(xy.Column) Then Columns.Add(xy.Column, New List(Of Control))
                    Columns(xy.Column).Add(Item)
                    If Not Rows.Keys.Contains(xy.Row) Then Rows.Add(xy.Row, New List(Of Control))
                    Rows(xy.Row).Add(Item)
                Next

                REM /// SIZING IS EASY IF ABSOLUTE- THE VALUE CAN BE TAKEN FROM THE ROWSTYLE Or COLUMNSTYLE VALUE AND IS ONLY THE CELLBORDERSTYLE WIDTH/HEIGHT + THE VALUE
                REM /// AUTOSIZING IS TRICKY- A MAX WIDTH OR HEIGHT VALUE MUST BE TAKEN FROM THE COLLECTION OF COLUMN Or ROW STRIPS

                Dim ColumnWidths As Integer = 0
                If Columns.Any Then
                    Dim ColumnStyles As New List(Of ColumnStyle)((From CS In TLP.ColumnStyles Select DirectCast(CS, ColumnStyle)))
                    Dim ColumnStyleAbsoluteWidths As Integer = (From CS In ColumnStyles Where CS.SizeType = SizeType.Absolute Select Convert.ToInt32(CS.Width + CellBorderThickness)).Sum
                    Dim ColumnStyleAutoWidths As Integer = (From CS In ColumnStyles Where CS.SizeType = SizeType.AutoSize Select Columns(ColumnStyles.IndexOf(CS)).Max(Function(c) c.Width) + CellBorderThickness).Sum
                    ColumnWidths = ColumnStyleAbsoluteWidths + ColumnStyleAutoWidths
                End If

                Dim RowHeights As Integer = 0
                If Rows.Any Then
                    Dim RowStyles As New List(Of RowStyle)((From RS In TLP.RowStyles Select DirectCast(RS, RowStyle)))
                    Dim RowStyleAbsoluteHeights As Integer = (From RS In RowStyles Where RS.SizeType = SizeType.Absolute Select Convert.ToInt32(RS.Height + CellBorderThickness)).Sum
                    Dim RowStyleAutoHeights As Integer = (From RS In RowStyles Where RS.SizeType = SizeType.AutoSize Select Rows(RowStyles.IndexOf(RS)).Max(Function(r) r.Height) + CellBorderThickness).Sum
                    RowHeights = RowStyleAbsoluteHeights + RowStyleAutoHeights
                End If

                Return New Size(BorderThickness + ColumnWidths, BorderThickness + RowHeights)
            Else
                Return Nothing
            End If

        End Function
        Public Sub Resize(TLP As TableLayoutPanel)
            If TLP IsNot Nothing Then TLP.Size = GetSize(TLP)
        End Sub
    End Module
End Namespace
Friend Class WindowWatch
    Friend Event Completed(sender As Object, e As IEnumerable(Of Process))
    Private WithEvents WindowTimer As New Timer With {.Interval = 200}
    Private ReadOnly Property WatchForText As String
    Private ReadOnly Property MaxTime As Long = 60000
    Friend ReadOnly Property Started As Date
    Friend ReadOnly Property Ended As Date
    Friend ReadOnly Property Succeeded As Boolean
    Friend ReadOnly Property Processes As List(Of Process)
    Friend Sub New(WindowText As String)
        _WatchForText = WindowText
    End Sub
    Friend Sub New(WindowText As String, MaxMiliseconds As Long)
        _WatchForText = WindowText
        _MaxTime = MaxMiliseconds
    End Sub
    Friend Sub Start()
        _Started = Now
        WindowTimer_Tick()
    End Sub
    Private Sub WindowTimer_Tick() Handles WindowTimer.Tick

        WindowTimer.Stop()
        Dim Windows = Process.GetProcesses
        Dim Window = From w In Windows Where w.MainWindowTitle.ToUpperInvariant.Contains(WatchForText.ToUpperInvariant) Select w
        If Window.Any Then
            _Processes = Window.ToList
            _Ended = Now
            _Succeeded = True
            RaiseEvent Completed(Me, Windows)
        Else
            'Keep waiting
            If (Now - Started).Milliseconds >= MaxTime Then
                _Ended = Now
                _Succeeded = False
                RaiseEvent Completed(Me, Windows)
            Else
                WindowTimer.Start()
            End If
        End If

    End Sub
End Class
Public NotInheritable Class SafeWalk
    Public Sub New()
    End Sub
    Public Shared Function EnumerateFiles(Path As String, SearchPattern As String, SearchOpt As SearchOption) As IEnumerable(Of String)

        Try
            Dim DirectoryFiles = Enumerable.Empty(Of String)()
            If SearchOpt = SearchOption.AllDirectories Then
                Try
                    DirectoryFiles = Directory.EnumerateDirectories(Path).SelectMany(Function(x) EnumerateFiles(x, SearchPattern, SearchOpt))
                    Return DirectoryFiles.Concat(Directory.EnumerateFiles(Path, SearchPattern))

                Catch ex As DirectoryNotFoundException
                    Return Enumerable.Empty(Of String)()

                End Try
            Else
                Return Enumerable.Empty(Of String)()

            End If
        Catch ex As UnauthorizedAccessException
            Return Enumerable.Empty(Of String)()
        End Try

    End Function
End Class
Public Module Ghost
#Region " MOUSE DOWN "
    Friend Enum SendInputEventType
        InputMouse
        InputKeyboard
        InputHardware
    End Enum
    <StructLayout(LayoutKind.Sequential)>
    Friend Structure INPUT
        Public type As SendInputEventType
        Public mkhi As MouseKeybdhardwareInputUnion
    End Structure
    <StructLayout(LayoutKind.Explicit)>
    Friend Structure MouseKeybdhardwareInputUnion
        <FieldOffset(0)>
        Friend mi As MouseInputData
        <FieldOffset(0)>
        Friend ki As KEYBDINPUT
        <FieldOffset(0)>
        Friend hi As HARDWAREINPUT
    End Structure
    <StructLayout(LayoutKind.Sequential)>
    Friend Structure KEYBDINPUT
        Friend wVk As UShort
        Friend wScan As UShort
        Friend dwFlags As UInteger
        Friend time As UInteger
        Friend dwExtraInfo As IntPtr
    End Structure
    <StructLayout(LayoutKind.Sequential)>
    Friend Structure HARDWAREINPUT
        Public uMsg As Integer
        Public wParamL As Short
        Public wParamH As Short
    End Structure
    Friend Structure MouseInputData
        Friend dx As Integer
        Friend dy As Integer
        Friend mouseData As UInteger
        Friend dwFlags As MouseEventFlags
        Friend time As UInteger
        Friend dwExtraInfo As IntPtr
    End Structure
    Friend Enum MouseEventFlags
        MOUSEEVENTFxMOVE = &H1
        MOUSEEVENTFxLEFTDOWN = &H2
        MOUSEEVENTFxLEFTUP = &H4
        MOUSEEVENTFxRIGHTDOWN = &H8
        MOUSEEVENTFxRIGHTUP = &H10
        MOUSEEVENTFxMIDDLEDOWN = &H20
        MOUSEEVENTFxMIDDLEUP = &H40
        MOUSEEVENTFxXDOWN = &H80
        MOUSEEVENTFxXUP = &H100
        MOUSEEVENTFxWHEEL = &H800
        MOUSEEVENTFxVIRTUALDESK = &H4000
        MOUSEEVENTFxABSOLUTE = &H8000
    End Enum
    Enum SystemMetric
        SMxCXSCREEN = 0
        SMxCYSCREEN = 1
    End Enum
    Private Function CalculateAbsoluteCoordinateX(ByVal x As Integer) As Integer
        Return CType((x * 65536) / NativeMethods.GetSystemMetrics(SystemMetric.SMxCXSCREEN), Integer)
    End Function
    Private Function CalculateAbsoluteCoordinateY(ByVal y As Integer) As Integer
        Return CType((y * 65536) / NativeMethods.GetSystemMetrics(SystemMetric.SMxCYSCREEN), Integer)
    End Function
    Public Sub ClickLeftMouseButton(ByVal Location As Point)
        ClickLeftMouseButton(Location.X, Location.Y)
    End Sub
    Public Sub ClickLeftMouseButton(ByVal x As Integer, ByVal y As Integer)

        Dim MouseInput As INPUT = New INPUT With {
            .type = SendInputEventType.InputMouse
        }
        With MouseInput
            .mkhi.mi.dx = CalculateAbsoluteCoordinateX(x)
            .mkhi.mi.dy = CalculateAbsoluteCoordinateY(y)
            .mkhi.mi.mouseData = 0
            .mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTFxMOVE Or MouseEventFlags.MOUSEEVENTFxABSOLUTE
            Dim unused1 = NativeMethods.SendInput(1, MouseInput, Marshal.SizeOf(New INPUT()))
            .mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTFxLEFTDOWN
            Dim unused2 = NativeMethods.SendInput(1, MouseInput, Marshal.SizeOf(New INPUT()))
            .mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTFxLEFTUP
            Dim unused3 = NativeMethods.SendInput(1, MouseInput, Marshal.SizeOf(New INPUT()))
        End With

    End Sub
    Public Sub MoveMouse(ByVal Location As Point)
        MoveMouse(Location.X, Location.Y)
    End Sub
    Public Sub MoveMouse(ByVal x As Integer, ByVal y As Integer)

        Dim MouseInput As INPUT = New INPUT With {
            .type = SendInputEventType.InputMouse
        }
        With MouseInput
            .mkhi.mi.dx = CalculateAbsoluteCoordinateX(x)
            .mkhi.mi.dy = CalculateAbsoluteCoordinateY(y)
            .mkhi.mi.mouseData = 0
            .mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTFxMOVE Or MouseEventFlags.MOUSEEVENTFxABSOLUTE
            Dim unused1 = NativeMethods.SendInput(1, MouseInput, Marshal.SizeOf(New INPUT()))
        End With

    End Sub
    Public Sub KeyPress(ByVal keyCode As Keys)

        Dim input As INPUT = New INPUT With {
        .type = SendInputEventType.InputKeyboard,
        .mkhi = New MouseKeybdhardwareInputUnion With {
            .ki = New KEYBDINPUT With {
                .wVk = CUShort(keyCode),
                .wScan = 0,
                .dwFlags = 0,
                .time = 0,
                .dwExtraInfo = IntPtr.Zero
            }
        }
    }
        Dim input2 As INPUT = New INPUT With {
        .type = SendInputEventType.InputKeyboard,
        .mkhi = New MouseKeybdhardwareInputUnion With {
            .ki = New KEYBDINPUT With {
                .wVk = CUShort(keyCode),
                .wScan = 0,
                .dwFlags = 2,
                .time = 0,
                .dwExtraInfo = IntPtr.Zero
            }
        }
    }
        Dim inputs As INPUT() = New INPUT() {input, input2}
        Dim unused1 = NativeMethods.SendInput(CUInt(inputs.Length), inputs, Marshal.SizeOf(GetType(INPUT)))

    End Sub
#End Region
End Module
Friend Module html
    Friend Event ElementWatched(sender As TimeSpan, e As List(Of HtmlElement))
    Private ReadOnly Property Document As HtmlDocument
    Private ReadOnly Property ElementIdName As String
    Private ReadOnly ElementStopWatch As New Stopwatch
    Private ReadOnly Property StopWatchLimit As Integer
    Private WithEvents ElementTimer As New Timer With {.Interval = 100}
    Friend TypePattern As String = "(?<= type=" & Chr(34) & ")[-a-z0-9_:.]{1,}(?=" & Chr(34) & ")"
    Friend idPattern As String = Replace(TypePattern, "type", "id")
    Friend namePattern As String = Replace(TypePattern, "type", "name")
    Friend Enum InputType
        'https://www.w3schools.com/html/html_form_input_types.asp
        '<input type ="button">
        '<input type="checkbox">
        '<input type="color">
        '<input type="date">
        '<input type="datetime-local">
        '<input type="email">
        '<input type="file">
        '<input type="hidden">
        '<input type="image">
        '<input type="month">
        '<input type="password">
        '<input type="radio">
        '<input type="range">
        '<input type="reset">
        '<input type="search">
        '<input type="submit">
        '<input type="tel">
        '<input type="text">
        '<input type="time">
        '<input type="url">
        '<input type="week">
        None
        Button
        Checkbox
        Color
        Email
        File
        Hidden
        Image
        Month
        Password
        Radio
        Range
        Reset
        Search
        Submit
        Tel
        Text
        Time
        Url
        Week
    End Enum
    Friend Enum SubmitType
        Click
        Enter
    End Enum
    Friend Function ElementSubmitType(Element As HtmlElement) As SubmitType

        Select Case ElementInputType(Element)
            Case InputType.Button, InputType.Submit
                Return SubmitType.Click
            Case InputType.Text, InputType.Password
                Return SubmitType.Enter
            Case Else
                Return SubmitType.Click
        End Select

    End Function
    Friend Function ElementInputType(Element As HtmlElement) As InputType

        Dim InputMatch As Match = Regex.Match(Element.OuterHtml, TypePattern, RegexOptions.IgnoreCase)
        If InputMatch.Success Then
            Return DirectCast([Enum].Parse(GetType(InputType), StrConv(InputMatch.Value, VbStrConv.ProperCase)), InputType)
        Else
            Return InputType.None
        End If

    End Function
    Friend Function ElementByRegex(Document As HtmlDocument, SearchValue As String) As HtmlElement

        Dim Elements As List(Of HtmlElement) = ElementsByRegex(Document, SearchValue)
        If Elements Is Nothing Then
            Return Nothing
        Else
            Return Elements.First
        End If

    End Function
    Friend Function ElementsByRegex(Document As HtmlDocument, SearchValue As String) As List(Of HtmlElement)

        If Document Is Nothing Then
            Return Nothing
        Else
            If Document.Body Is Nothing Then
                Return Nothing
            Else
                Dim OuterHtml As String = Document.Body.OuterHtml
                Dim ids As New List(Of Match)(From r In Regex.Matches(OuterHtml, idPattern, RegexOptions.IgnoreCase) Select DirectCast(r, Match))
                Dim names As New List(Of Match)(From r In Regex.Matches(OuterHtml, namePattern, RegexOptions.IgnoreCase) Select DirectCast(r, Match))
                Dim MatchingIds = From i In ids.Union(names) Where Regex.Match(i.Value, SearchValue, RegexOptions.IgnoreCase).Success
                If MatchingIds.Any Then
                    Dim Elements = New List(Of HtmlElement)(From html In MatchingIds Select Document.GetElementById(html.Value))
                    Return Elements.Distinct.ToList
                Else
                    Return Nothing
                End If
            End If
        End If

    End Function
    Friend Function ElementsByTag(Document As HtmlDocument) As Dictionary(Of String, List(Of HtmlElement))

        Dim Elements As New Dictionary(Of String, List(Of HtmlElement))
        Dim OuterHtml As String = Document.Body.OuterHtml
        Dim Tags = New List(Of String) From {"a", "body", "br", "div", "Form", "h1", "h2", "h3", "h4", "head", "html", "iframe", "img", "input", "li", "link", "meta", "ol", "OptionOn", "p", "script", "select", "span", "style", "table", "th", "td", "textarea", "title", "tr", "ul"}
        For Each Tag In Tags
            For Each Element As HtmlElement In Document.GetElementsByTagName(Tag)
                If Not Elements.ContainsKey(Tag) Then Elements.Add(Tag, New List(Of HtmlElement))
                Elements(Tag).Add(Element)
            Next
        Next
        Return Elements

    End Function
    Friend Function ElementsAll(Document As HtmlDocument) As List(Of HtmlElement)

        Dim ElementsDictionary = ElementsByTag(Document)
        Dim AllElements As New List(Of HtmlElement)
        For Each Tag In ElementsDictionary.Keys
            AllElements.AddRange(ElementsDictionary(Tag))
        Next
        Return AllElements

    End Function
    Friend Function ElementsByKeyText(Document As HtmlDocument, SearchValue As String) As List(Of HtmlElement)

        If Document Is Nothing Then
            Return Nothing
        Else
            If Document.Body Is Nothing Then
                Return Nothing
            Else
                SearchValue = Replace(SearchValue, "`", Chr(34))
                Dim MatchingElements = From e In ElementsAll(Document) Where e.InnerHtml IsNot Nothing AndAlso Regex.Match(e.InnerHtml, SearchValue, RegexOptions.IgnoreCase).Success
                Return MatchingElements.ToList
            End If
        End If

    End Function
    Friend Function SubmitForm(Document As HtmlDocument, Element As HtmlElement) As HtmlElement

        If Document Is Nothing Or Element Is Nothing Then
            Return Nothing
        Else
            With Document
                Dim Forms As HtmlElementCollection = .GetElementsByTagName("form")
                Dim FormsBounds = From f In Forms Where DirectCast(f, HtmlElement).OffsetRectangle.Contains(Element.OffsetRectangle) Select DirectCast(f, HtmlElement)
                If FormsBounds.Any Then
                    Dim InnerMostForm = FormsBounds.Where(Function(f) f.OffsetRectangle.Width * f.OffsetRectangle.Height = FormsBounds.Min(Function(fb) fb.OffsetRectangle.Width * fb.OffsetRectangle.Height))
                    If InnerMostForm.Any Then
                        Return InnerMostForm.First
                    Else
                        Return Nothing
                    End If
                Else
                    Return Nothing
                End If
            End With
        End If

    End Function
    Friend Sub ElementWatch(WebDocument As HtmlDocument, IdName As String, Optional Timeout As Integer = 10)
        _Document = WebDocument
        _ElementIdName = IdName
        _StopWatchLimit = Timeout
        ElementStopWatch.Start()
        ElementTimer.Start()
    End Sub
    Private Sub ElementTimer_Tick() Handles ElementTimer.Tick

        ElementTimer.Stop()
        If ElementStopWatch.Elapsed.TotalSeconds < StopWatchLimit Then
            Dim Elements As New List(Of HtmlElement)(ElementsByKeyText(Document, ElementIdName))
            If Elements Is Nothing Then
                'Still looking
                ElementTimer.Start()
            Else
                If Elements.Any Then
                    'Succeeded
                    ElementStopWatch.Stop()
                    RaiseEvent ElementWatched(ElementStopWatch.Elapsed, Elements)
                Else
                    'Still looking
                    ElementTimer.Start()
                End If
            End If
        Else
            'Failed
            ElementStopWatch.Stop()
            RaiseEvent ElementWatched(ElementStopWatch.Elapsed, Nothing)
        End If

    End Sub
End Module
Friend Class PropertyConverter
    Inherits TypeConverter
    Public Overloads Overrides Function CanConvertFrom(ByVal context As ITypeDescriptorContext, ByVal sourceType As Type) As Boolean
        If (sourceType.Equals(GetType(String))) Then
            Return True
        Else
            Return MyBase.CanConvertFrom(context, sourceType)
        End If
    End Function
    Public Overloads Overrides Function CanConvertTo(ByVal context As ITypeDescriptorContext, ByVal destinationType As Type) As Boolean
        If (destinationType.Equals(GetType(String))) Then
            Return True
        Else
            Return MyBase.CanConvertTo(context, destinationType)
        End If
    End Function
    Public Overloads Overrides Function ConvertTo(ByVal context As ITypeDescriptorContext, ByVal culture As Globalization.CultureInfo, ByVal value As Object, ByVal destinationType As Type) As Object
        If (destinationType.Equals(GetType(String))) Then
            Return value.ToString()
        Else
            Return MyBase.ConvertTo(context, culture, value, destinationType)
        End If
    End Function
    Public Overloads Overrides Function GetPropertiesSupported(ByVal context As ITypeDescriptorContext) As Boolean
        Return True
    End Function
    Public Overloads Overrides Function GetProperties(ByVal context As ITypeDescriptorContext, ByVal value As Object, ByVal Attribute() As Attribute) As PropertyDescriptorCollection
        Return TypeDescriptor.GetProperties(value)
    End Function
End Class
Public Class CursorBusy
    Implements IDisposable
#Region " DISPOSE "
    Dim disposed As Boolean = False
    ReadOnly Handle As SafeHandle = New Microsoft.Win32.SafeHandles.SafeFileHandle(IntPtr.Zero, True)
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Protected Overridable Sub Dispose(disposing As Boolean)
        If disposed Then Return
        If disposing Then
            Handle.Dispose()
            ' Free any other managed objects here.
            Cursor.Current.Dispose()
        End If
        disposed = True
    End Sub
#End Region
    Public Sub New()
        Cursor.Current = Cursors.WaitCursor
    End Sub
End Class
Public NotInheritable Class CursorHelper
    Private Structure IconInfo
        Public fIcon As Boolean
        Public xHotspot As Integer
        Public yHotspot As Integer
        Public hbmMask As IntPtr
        Public hbmColor As IntPtr
    End Structure
    Public Shared Function CreateCursor(ByVal bmp As Bitmap, ByVal xHotspot As Integer, ByVal yHotspot As Integer) As Cursor

        If bmp Is Nothing Then
            Return Cursors.Default
        Else
            Dim tmp As New IconInfo With {
                .xHotspot = xHotspot,
                .yHotspot = yHotspot,
                .fIcon = False,
                .hbmMask = bmp.GetHbitmap(),
                .hbmColor = bmp.GetHbitmap()
            }
            Dim Pointer As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(tmp))
            Marshal.StructureToPtr(tmp, Pointer, True)
            Dim CursorPointer As IntPtr = NativeMethods.CreateIconIndirect(Pointer)
            NativeMethods.DestroyIcon(Pointer)
            NativeMethods.DeleteObject(tmp.hbmMask)
            NativeMethods.DeleteObject(tmp.hbmColor)
            Return New Cursor(CursorPointer)
        End If

    End Function
End Class
Public NotInheritable Class WebBrowserUpdater
    Friend Shared ReadOnly is64BitProcess As Boolean = (IntPtr.Size = 8)
    Friend Shared ReadOnly is64BitOperatingSystem As Boolean = is64BitProcess OrElse InternalCheckIsWow64()
    Public Shared Function InternalCheckIsWow64() As Boolean
        If (Environment.OSVersion.Version.Major = 5 AndAlso Environment.OSVersion.Version.Minor >= 1) OrElse Environment.OSVersion.Version.Major >= 6 Then
            Using p As Process = Process.GetCurrentProcess()
                Dim retVal As Boolean
                If Not NativeMethods.IsWow64Process(p.Handle, retVal) Then
                    Return False
                End If
                Return retVal
            End Using
        Else
            Return False
        End If
    End Function
    Public Shared Function GetEmbVersion() As Integer
        Dim ieVer As Integer = GetBrowserVersion()

        If ieVer > 9 Then
            Return ieVer * 1000 + 1
        End If

        If ieVer > 7 Then
            Return ieVer * 1111
        End If

        Return 7000
    End Function
    ' End Function GetEmbVersion
    Public Shared Sub FixBrowserVersion()
        Dim appName As String = Path.GetFileNameWithoutExtension(Reflection.Assembly.GetExecutingAssembly().Location)
        FixBrowserVersion(appName)
    End Sub

    Public Shared Sub FixBrowserVersion(ByVal appName As String)
        FixBrowserVersion(appName, GetEmbVersion())
    End Sub
    ' End Sub FixBrowserVersion
    Public Shared Sub FixBrowserVersion(ByVal appName As String, ByVal ieVer As Integer)
        FixBrowserVersion_Internal("HKEY_LOCAL_MACHINE", appName & ".exe".ToString(InvariantCulture), ieVer)
        FixBrowserVersion_Internal("HKEY_CURRENT_USER", appName & ".exe".ToString(InvariantCulture), ieVer)
        FixBrowserVersion_Internal("HKEY_LOCAL_MACHINE", appName & ".vshost.exe".ToString(InvariantCulture), ieVer)
        FixBrowserVersion_Internal("HKEY_CURRENT_USER", appName & ".vshost.exe".ToString(InvariantCulture), ieVer)
    End Sub
    ' End Sub FixBrowserVersion
    Private Shared Sub FixBrowserVersion_Internal(ByVal root As String, ByVal appName As String, ByVal ieVer As Integer)
        Try
            'For 64 bit Machine 
            If InternalCheckIsWow64() Then
                Microsoft.Win32.Registry.SetValue(root & "\Software\Wow6432Node\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION".ToString(InvariantCulture), appName, ieVer)
            Else
                'For 32 bit Machine 
                Microsoft.Win32.Registry.SetValue(root & "\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION".ToString(InvariantCulture), appName, ieVer)

            End If
        Catch generatedExceptionName As ArgumentNullException
            Dim MessageBody As String = "You have to be administrator to run start this process. Please close the software. Right click on the iGiftCard icon and select RUN AS ADMINISTRATOR.".ToString(InvariantCulture)
            MessageBox.Show(MessageBody, "Administrator".ToString(InvariantCulture), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub
    ' End Sub FixBrowserVersion_Internal
    Public Shared Function GetBrowserVersion() As Integer

        Dim strKeyPath As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer"
        Dim ls As String() = New String() {"svcVersion", "svcUpdateVersion", "Version", "W2kVersion"}

        Dim maxVer As Integer = 0
        For i As Integer = 0 To ls.Length - 1
            Dim objVal As Object = Microsoft.Win32.Registry.GetValue(strKeyPath, ls(i), "0")
            Dim strVal As String = System.Convert.ToString(objVal, InvariantCulture)
            If strVal IsNot Nothing Then
                Dim iPos As Integer = strVal.IndexOf("."c)
                If iPos > 0 Then
                    strVal = strVal.Substring(0, iPos)
                End If

                Dim res As Integer = 0
                If Integer.TryParse(strVal, res) Then
                    maxVer = Math.Max(maxVer, res)
                End If
            End If
        Next
        Return maxVer

    End Function
    ' End Function GetBrowserVersion
End Class
Public Class VerticalScrollBar
    Friend WithEvents Timer As New Timer With {.Interval = 250}
    Friend Alpha As Byte = 128, UpAlpha As Byte, DownAlpha As Byte
    Private Const Width As Integer = 12, ArrowsHeight As Integer = Width + 2, ShadowDepth As Integer = 8
#Region " Constructor "
    Public Sub New(Control As Control)

        Me.Control = Control
        If Control IsNot Nothing Then
            AddHandler Control.SizeChanged, AddressOf ControlSizeChanged
            AddHandler Control.MouseDown, AddressOf MouseDown
            AddHandler Control.MouseMove, AddressOf MouseMove
            AddHandler Control.MouseUp, AddressOf MouseUp
            AddHandler Control.MouseHover, AddressOf MouseHeld
        End If

    End Sub
#End Region
#Region " Properties and Fields"
    Private mScrolling As Boolean
    Friend ReadOnly Property Scrolling As Boolean
        Get
            Return mScrolling
        End Get
    End Property
    Private mReference As New Point
    Friend Property Reference As Point
        Get
            Return mReference
        End Get
        Set(value As Point)
            If Not mReference = value Then
                mReference = value
            End If
        End Set
    End Property
    Public Property Lines As Boolean
    Public Property Color As Color = Color.CornflowerBlue
    Public Property SmallChange As Integer = 1
    Public Property LargeChange As Integer
    Public ReadOnly Property Control As Control
    Friend ReadOnly Property Pages As List(Of Double)
        Get
            Dim PageCount As Double = ScrollHeight / Bounds.Height
            Return Enumerable.Range(0, Convert.ToInt32(Math.Floor(PageCount))).Select(Function(x) ArrowsHeight + (x * Bounds.Height) / 2).ToList
        End Get
    End Property
    Private _Value As Integer
    Public Property Value As Integer
        Get
            Return _Value
        End Get
        Set(value As Integer)
            If Not (value = _Value) Then
                If value < 0 Then
                    _Value = 0
                ElseIf (value) > ScrollHeight Then
                    _Value = ScrollHeight
                Else
                    _Value = value
                End If
                RaiseEvent ValueChanged(Me, Nothing)
            End If
        End Set
    End Property
    Private _Height As Integer
    Public Property Height As Integer
        Get
            Return _Height
        End Get
        Set(value As Integer)
            UpdateBounds()
            _Height = value
        End Set
    End Property
    Private _Maximum As Integer
    Public Property Maximum As Integer
        Get
            Return _Maximum
        End Get
        Set(value As Integer)
            UpdateBounds()
            _Maximum = value
        End Set
    End Property
    Friend ReadOnly Property ScrollHeight As Integer
        Get
            Return (Maximum - Height)
        End Get
    End Property
    Private _Bounds As New Rectangle(0, 0, Width, 0)
    Public ReadOnly Property Bounds As Rectangle
        Get
            Return _Bounds
        End Get
    End Property
    Private _TrackBounds As New Rectangle(0, ArrowsHeight, Width, 0)
    Public ReadOnly Property TrackBounds As Rectangle
        Get
            Return _TrackBounds
        End Get
    End Property
    Private _UpBounds As New Rectangle(0, -1, Width, ArrowsHeight)
    Friend ReadOnly Property UpBounds As Rectangle
        Get
            Return _UpBounds
        End Get
    End Property
    Private _BarBounds As New Rectangle(0, ArrowsHeight, Width, 0)
    Friend ReadOnly Property BarBounds As Rectangle
        Get
            If _BarBounds.Top <= UpBounds.Bottom Then
                _BarBounds.Y = UpBounds.Bottom
            ElseIf _BarBounds.Bottom >= DownBounds.Top Then
                _BarBounds.Y = (DownBounds.Top - _BarBounds.Height)
            End If
            Return _BarBounds
        End Get
    End Property
    Private _DownBounds As New Rectangle(0, 0, Width, ArrowsHeight)
    Friend ReadOnly Property DownBounds As Rectangle
        Get
            Return _DownBounds
        End Get
    End Property
    Friend ReadOnly Property Visible As Boolean
        Get
            Return (ScrollHeight > Height)
        End Get
    End Property
#End Region
#Region " Events "
    Public Event ValueChanged(ByVal sender As Object, e As EventArgs)
    Private Sub ControlSizeChanged(ByVal sender As Object, e As EventArgs)
        UpdateBounds()
    End Sub
    Private Sub MouseDown(ByVal sender As Object, e As MouseEventArgs)
        If UpBounds.Contains(e.Location) Then
            Reference = e.Location
            Timer.Start()
            Value -= SmallChange
        ElseIf DownBounds.Contains(e.Location) Then
            Reference = e.Location
            Timer.Start()
            Value += SmallChange
        ElseIf TrackBounds.Contains(e.Location) Then
            If Not BarBounds.Contains(e.Location) Then
                Dim TrackValue As Double = ((e.Y - TrackBounds.Top) / (TrackBounds.Height - BarBounds.Height) * ScrollHeight)
                Value = Convert.ToInt32(Math.Floor(TrackValue / SmallChange) * SmallChange)
                _BarBounds.Y = e.Y
            End If
            Reference = e.Location
            Alpha = 255
            Control.Invalidate()
        End If
        Reference = e.Location
    End Sub
    Private Sub MouseHeld(ByVal sender As Object, e As EventArgs) Handles Timer.Tick
        If UpBounds.Contains(Reference) Then
            Value -= LargeChange
            _BarBounds.Y = Convert.ToInt32(Value * (TrackBounds.Height - BarBounds.Height) / ScrollHeight) + TrackBounds.Top
        ElseIf DownBounds.Contains(Reference) Then
            Value += LargeChange
            _BarBounds.Y = Convert.ToInt32(Value * (TrackBounds.Height - BarBounds.Height) / ScrollHeight) + TrackBounds.Top
        End If
        Control.Invalidate()
    End Sub
    Private Sub MouseMove(ByVal sender As Object, e As MouseEventArgs)
        Alpha = 60
        UpAlpha = 0
        DownAlpha = 0
        If Bounds.Contains(e.Location) Or Scrolling Then
            If e.Y < TrackBounds.Top Or e.Y > TrackBounds.Bottom Then
                mScrolling = False
            End If
            If TrackBounds.Contains(e.Location) Or Scrolling Then
                Timer.Stop()
                If e.Button = MouseButtons.Left Then
                    Alpha = 255
                    mScrolling = True
                    Dim Change As Integer = (e.Y - Reference.Y)
                    _BarBounds.Y += Change
                    Dim TrackValue As Double = ((BarBounds.Top - TrackBounds.Top) / (TrackBounds.Height - BarBounds.Height) * ScrollHeight)
                    Value = Convert.ToInt32(Math.Floor(TrackValue / SmallChange) * SmallChange)
                    Reference = e.Location
                Else
                    mScrolling = False
                    If BarBounds.Contains(e.Location) Then Alpha = 128
                End If
            ElseIf UpBounds.Contains(e.Location) Then
                UpAlpha = 64
            ElseIf DownBounds.Contains(e.Location) Then
                DownAlpha = 64
            End If
            Control.Invalidate()
        End If
    End Sub
    Private Sub MouseUp(ByVal sender As Object, e As MouseEventArgs)
        If Bounds.Contains(e.Location) Then
            Reference = Nothing
        End If
        mScrolling = False
        Control.Invalidate()
    End Sub
#End Region
#Region " Methods "
    Private Sub UpdateBounds()

        With _Bounds
            .X = Control.Width - Width - ShadowDepth
            .Height = Control.Height - ShadowDepth
            .Width = If(Visible, Width, 0)
        End With
        With _TrackBounds
            .X = _Bounds.X
            .Height = _Bounds.Height - (ArrowsHeight * 2)
            .Width = _Bounds.Width
        End With
        With _BarBounds
            .X = _Bounds.X - 1
            .Width = _Bounds.Width
            .Height = If(Visible, {Convert.ToInt32((Height / Maximum) * TrackBounds.Height), 20}.Max, 0)
        End With
        With _UpBounds
            .X = _Bounds.X - 1
            .Width = _Bounds.Width
        End With
        With _DownBounds
            .X = _Bounds.X - 1
            .Y = _Bounds.Height - ArrowsHeight
            .Width = _Bounds.Width
        End With

    End Sub
#End Region
End Class
Public NotInheritable Class CustomRenderer
    Inherits ToolStripProfessionalRenderer
    Public Enum ColorTheme
        Brown
        Green
        Blue
        Red
        Gray
        Yellow
    End Enum
    Public Property Theme As ColorTheme
    Protected Overrides Sub OnRenderImageMargin(ByVal e As ToolStripRenderEventArgs)

        MyBase.OnRenderImageMargin(e)
        If e IsNot Nothing Then
            Dim MarginWidth As Integer = e.AffectedBounds.Width
            e.ToolStrip.Items.OfType(Of ToolStripControlHost)().ToList().ForEach(Function(Item) As ToolStripControlHost
                                                                                     If IsNothing(Item.Image) Then
                                                                                     Else
                                                                                         Dim Size = Item.GetCurrentParent().ImageScalingSize
                                                                                         Dim Location = Item.Bounds.Location
                                                                                         Dim OffsetX As Integer = Convert.ToInt32((MarginWidth - Item.Image.Width) / 2)
                                                                                         Dim OffsetY As Integer = Convert.ToInt32((Item.Height - Item.Image.Height) / 2)
                                                                                         Location = New Point(OffsetX, Location.Y + OffsetY)
                                                                                         Dim ImageRectangle = New Rectangle(Location, Size)
                                                                                         e.Graphics.DrawImage(Item.Image,
                                                                                                              ImageRectangle,
                                                                                                              New Rectangle(Point.Empty, Item.Image.Size),
                                                                                                              GraphicsUnit.Pixel)
                                                                                     End If
                                                                                     Return Nothing
                                                                                 End Function)
        End If

    End Sub
    Protected Overrides Sub OnRenderMenuItemBackground(ByVal e As ToolStripItemRenderEventArgs)

        If e IsNot Nothing Then
            If e.Item.Selected Then
                Using Brush As New Drawing2D.LinearGradientBrush(e.Item.ContentRectangle, Color.FromArgb(255, 227, 224, 215), Color.White, Drawing2D.LinearGradientMode.Vertical)
                    e.Graphics.FillRectangle(Brush, e.Item.ContentRectangle)
                End Using
                Dim RoundRectangle As Rectangle = e.Item.ContentRectangle
                RoundRectangle.Inflate(-2, -2)
                RoundRectangle.Offset(0, -2)
                Using GP As Drawing2D.GraphicsPath = DrawRoundedRectangle(e.Item.ContentRectangle)
                    Using PathPen As New Pen(Color.Peru, 1)
                        e.Graphics.DrawPath(PathPen, GP)
                    End Using
                End Using
            Else
                Using Brush As New SolidBrush(Color.FromArgb(255, 227, 224, 215))
                    e.Graphics.FillRectangle(Brush, e.Item.ContentRectangle)
                End Using
            End If
        End If

    End Sub
End Class
Public Module ThreadHelperClass
    'Delegate Sub SetTextCallback(ByVal f As Form, ByVal ctrl As Control, ByVal text As String)
    'Delegate Sub SetControlTextCallback(ByVal ctrl As Control, ByVal text As String)
    'Delegate Sub SetIconCallback(ByVal f As Form, ByVal icon As Icon)
    Delegate Sub SetPropertyCallback(ByVal c As Control, ByVal n As String, v As Object)
    'Public Sub SetSafeIcon(ByVal form As Form, ByVal icon As Icon)

    '    If form Is Nothing Then
    '    Else
    '        If form.InvokeRequired Then
    '            Dim d As SetIconCallback = New SetIconCallback(AddressOf SetSafeIcon)
    '            form.Invoke(d, New Object() {form, icon})
    '        Else
    '            form.Icon = icon
    '        End If
    '    End If

    'End Sub
    Public Sub SetSafeControlPropertyValue(ByVal Item As Control, ByVal PropertyName As String, PropertyValue As Object)

        If Item Is Nothing Then
        Else
            Dim t As Type = Item.GetType
            If Item.InvokeRequired Then
                Dim d As SetPropertyCallback = New SetPropertyCallback(AddressOf SetSafeControlPropertyValue)
                Item.Invoke(d, New Object() {Item, PropertyName, PropertyValue})
            Else
                Dim pi As PropertyInfo = t.GetProperty(PropertyName)
                pi.SetValue(Item, PropertyValue)
            End If
        End If

    End Sub
    'Public Sub SetSafeText(ByVal form As Form, ByVal text As String)
    '    SetSafeText(form, form, text)
    'End Sub
    'Public Sub SetSafeText(ByVal form As Form, ByVal ctrl As Control, ByVal text As String)

    '    If form Is Nothing Or ctrl Is Nothing Then
    '    Else
    '        If ctrl.InvokeRequired Then
    '            Dim d As SetTextCallback = New SetTextCallback(AddressOf SetSafeText)
    '            form.Invoke(d, New Object() {form, ctrl, text})
    '        Else
    '            ctrl.Text = text
    '        End If
    '    End If

    'End Sub
    'Public Sub SetSafeControlText(ByVal ctrl As Control, ByVal text As String)

    '    If ctrl IsNot Nothing Then
    '        If ctrl.InvokeRequired Then
    '            Dim d As SetControlTextCallback = New SetControlTextCallback(AddressOf SetSafeControlText)
    '            ctrl.Invoke(d, New Object() {ctrl, text})
    '        Else
    '            ctrl.Text = text
    '        End If
    '    End If

    'End Sub
End Module
Public NotInheritable Class AlertEventArgs
    Inherits EventArgs
    Public ReadOnly Property Message As String
    Public Sub New(Value As String)
        Message = Value
    End Sub
End Class

#Region " DLLs "
<StructLayout(LayoutKind.Sequential)>
Public Structure SCROLLINFO
    Implements IEquatable(Of SCROLLINFO)
    Public Property CbSize As Integer
    Public Property FMask As Integer
    Public ReadOnly Property NMin As Integer
    Public ReadOnly Property NMax As Integer
    Public ReadOnly Property NPage As Integer
    Public ReadOnly Property NPos As Integer
    Public ReadOnly Property NTrackPos As Integer
    Public Overrides Function ToString() As String
        Return Join({"Size=" + CbSize.ToString(InvariantCulture),
                    "Mask=" + FMask.ToString(InvariantCulture),
                    "Min=" + NMin.ToString(InvariantCulture),
                    "Max=" + NMax.ToString(InvariantCulture),
                    "Page=" + NPage.ToString(InvariantCulture),
                    "Pos=" + NPos.ToString(InvariantCulture),
                    "Track=" + NTrackPos.ToString(InvariantCulture)}, ",")
    End Function
    Public Overrides Function GetHashCode() As Integer
        Return CbSize.GetHashCode Xor FMask.GetHashCode Xor NMin.GetHashCode Xor NPage.GetHashCode Xor NPos.GetHashCode Xor NTrackPos.GetHashCode
    End Function
    Public Overloads Function Equals(ByVal other As SCROLLINFO) As Boolean Implements IEquatable(Of SCROLLINFO).Equals
        Return CbSize = other.CbSize AndAlso FMask = other.FMask AndAlso NMin = other.NMin
    End Function
    Public Shared Operator =(ByVal value1 As SCROLLINFO, ByVal value2 As SCROLLINFO) As Boolean
        Return value1.Equals(value2)
    End Operator
    Public Shared Operator <>(ByVal value1 As SCROLLINFO, ByVal value2 As SCROLLINFO) As Boolean
        Return Not value1 = value2
    End Operator
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is SCROLLINFO Then
            Return CType(obj, SCROLLINFO) = Me
        Else
            Return False
        End If
    End Function
End Structure
Friend NotInheritable Class NativeMethods
    Private Sub New()
    End Sub
    <DllImport("user32.dll", EntryPoint:="GetScrollInfo")>
    Friend Shared Function GetScrollInfo(ByVal hwnd As IntPtr, ByVal nBar As Integer, ByRef lpsi As SCROLLINFO) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Friend Shared Function GetScrollPos(ByVal hWnd As IntPtr, ByVal nBar As Integer) As Integer
    End Function
    <DllImport("user32.dll")>
    Friend Shared Function SetScrollPos(ByVal hWnd As IntPtr, ByVal nBar As Integer, ByVal nPos As Integer, ByVal bRedraw As Boolean) As Integer
    End Function
    <DllImport("user32.dll")>
    Friend Shared Function GetCursorPos(ByRef lpPoint As Point) As Boolean
    End Function
    <DllImport("user32.dll")>
    Friend Shared Function SetCursorPos(ByVal x As Integer, ByVal Y As Integer) As Boolean
    End Function
    <DllImport("User32.dll", CharSet:=CharSet.Auto)>
    Friend Shared Function ReleaseDC(hWnd As IntPtr, hDC As IntPtr) As Integer
    End Function
    <DllImport("User32.dll")>
    Friend Shared Function GetWindowDC(hWnd As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Friend Shared Function SendInput(ByVal nInputs As UInteger, ByRef pInputs As INPUT, ByVal cbSize As Integer) As UInteger
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Friend Shared Function SendInput(ByVal numberOfInputs As UInteger, ByVal inputs As INPUT(), ByVal sizeOfInputStructure As Integer) As UInteger
    End Function
    <DllImport("kernel32.dll", CallingConvention:=CallingConvention.Winapi, SetLastError:=True)>
    Friend Shared Function IsWow64Process(<[In]()> ByVal hProcess As IntPtr, <Out()> ByRef wow64Process As Boolean) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Friend Shared Function MessageBeep(ByVal uType As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function
    <DllImport("user32.dll", EntryPoint:="CreateIconIndirect")>
    Friend Shared Function CreateIconIndirect(ByVal iconInfo As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Friend Shared Function DestroyIcon(ByVal handle As IntPtr) As Boolean
    End Function
    <DllImport("gdi32.dll")>
    Friend Shared Function DeleteObject(ByVal hObject As IntPtr) As Boolean
    End Function
    Friend Declare Auto Function GetSystemMetrics Lib "user32.dll" (ByVal smIndex As Integer) As Integer
    Friend Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Integer) As Short
    Friend Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As IntPtr) As Integer
End Class
#End Region

Class Module1
    Public Shared Sub Main()
        ' This variable holds the amount of indenting that 
        ' should be used when displaying each line of information.
        Dim indent As Integer = 0
        ' Display information about the EXE assembly.
        Dim a As Assembly = GetType(Module1).Assembly
        Display(indent, "Assembly identity={0}", a.FullName)
        Display((indent + 1), "Codebase={0}", a.CodeBase)
        ' Display the set of assemblies our assemblies reference.
        Display(indent, "Referenced assemblies:")
        For Each an As AssemblyName In a.GetReferencedAssemblies
            Display((indent + 1), "Name={0}, Version={1}, Culture={2}, PublicKey token={3}", an.Name, an.Version, an.CultureInfo.Name, BitConverter.ToString(an.GetPublicKeyToken))
        Next
        Display(indent, "")
        ' Display information about each assembly loading into this AppDomain.
        For Each b As Assembly In AppDomain.CurrentDomain.GetAssemblies
            Display(indent, "Assembly: {0}", b)
            ' Display information about each module of this assembly.
            For Each m As [Module] In b.GetModules(True)
                Display((indent + 1), "Module: {0}", m.Name)
            Next
            ' Display information about each type exported from this assembly.
            indent = (indent + 1)
            For Each t As Type In b.GetExportedTypes
                Display(0, "")
                Display(indent, "Type: {0}", t)
                ' For each type, show its members & their custom attributes.
                indent = (indent + 1)
                For Each mi As MemberInfo In t.GetMembers
                    Display(indent, "Member: {0}", mi.Name)
                    DisplayAttributes(indent, mi)
                    ' If the member is a method, display information about its parameters.
                    If (mi.MemberType = MemberTypes.Method) Then
                        For Each pi As ParameterInfo In CType(mi, MethodInfo).GetParameters
                            Display((indent + 1), "Parameter: Type={0}, Name={1}", pi.ParameterType, pi.Name)
                        Next
                    End If

                    ' If the member is a property, display information about the property's accessor methods.
                    If (mi.MemberType = MemberTypes.Property) Then
                        For Each am As MethodInfo In CType(mi, PropertyInfo).GetAccessors
                            Display((indent + 1), "Accessor method: {0}", am)
                        Next
                    End If

                Next
                indent = (indent - 1)
            Next
            indent = (indent - 1)
        Next
    End Sub
    ' Displays the custom attributes applied to the specified member.
    Public Shared Sub DisplayAttributes(ByVal indent As Integer, ByVal mi As MemberInfo)
        ' Get the set of custom attributes; if none exist, just return.
        Dim attrs() As Object = mi.GetCustomAttributes(False)
        If (attrs.Length = 0) Then
            Return
        End If

        ' Display the custom attributes applied to this member.
        Display((indent + 1), "Attributes:")
        For Each o As Object In attrs
            Display((indent + 2), "{0}", o.ToString)
        Next
    End Sub
    ' Display a formatted string indented by the specified amount.
    Public Shared Sub Display(ByVal indent As Integer, ByVal format As String, ParamArray ByVal param() As Object)
        Console.Write(New String(ChrW(32), (indent * 2)))
        Console.WriteLine(format, param)
    End Sub
End Class

Class CustomToolTip
    Inherits ToolTip
    Public Sub New()
        MyBase.New
        OwnerDraw = True
        IsBalloon = True
        AddHandler Popup, AddressOf OnPopup
        AddHandler Draw, AddressOf OnDraw
    End Sub
    Private Sub OnPopup(ByVal sender As Object, ByVal e As PopupEventArgs)
        e.ToolTipSize = New Size(200, 100)
    End Sub
    Private Sub OnDraw(ByVal sender As Object, ByVal e As DrawToolTipEventArgs)

        Using g As Graphics = e.Graphics
            Using b As Drawing2D.LinearGradientBrush = New Drawing2D.LinearGradientBrush(e.Bounds, Color.GreenYellow, Color.MintCream, 45.0!)
                g.FillRectangle(Brushes.Transparent, e.Bounds)

                Dim point1 As PointF = New PointF(0.0F, 100.0F)
                Dim point2 As PointF = New PointF(200.0F, 50.0F)
                Dim point3 As PointF = New PointF(250.0F, 200.0F)
                Dim point4 As PointF = New PointF(50.0F, 150.0F)
                Dim points() = {point1, point2, point3, point4}
                g.FillClosedCurve(b, points)

                Using Pen1 As New Pen(Brushes.Red, 1)
                    g.DrawRectangle(Pen1, New Rectangle(e.Bounds.X, e.Bounds.Y, (e.Bounds.Width - 1), (e.Bounds.Height - 1)))
                End Using
                g.DrawString(e.ToolTipText, e.Font, Brushes.Silver, New PointF((e.Bounds.X + 6), (e.Bounds.Y + 6)))
                ' shadow layer
                g.DrawString(e.ToolTipText, e.Font, Brushes.Black, New PointF((e.Bounds.X + 5), (e.Bounds.Y + 5)))
                ' top layer
            End Using
        End Using

    End Sub
End Class