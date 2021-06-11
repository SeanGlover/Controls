Option Explicit On
Option Strict On
Imports System.Text
Imports System.IO
Imports System.IO.Compression
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Text.RegularExpressions
Imports System.Data.OleDb
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Globalization
Imports ExcelDataReader

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
    Public Const ArrayDelimiter As String = " º "
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
    Public Const FilePattern As String = "^[A-Z]:(\\[^\/\\:*<>|]{1,}){1,}\.[a-z]{3,4}"
    Public Const SelectPattern As String = "SELECT[^■]{1,}?(?=FROM)"
    Public Const ObjectPattern As String = "([A-Z0-9!%{}^~_@#$]{1,}([.][A-Z0-9!%{}^~_@#$]{1,}){0,2})"     'DataSource.Owner.Name
    Public Const FieldPattern As String = "[\s]{1,}\([A-Z0-9!%{}^~_@#$]{1,}(,[\s]{0,}[A-Z0-9!%{}^~_@#$]{1,}){0,}\)"
    Public Const FromJoinCommaPattern As String = "(?<=FROM |JOIN )[\s]{0,}[A-Z0-9!%{}^~_@#$♥]{1,}([.][A-Z0-9!%{}^~_@#$♥]{1,}){0,2}([\s]{1,}[A-Z0-9!%{}^~_@#$]{1,}){0,1}|(?<=,)[\s]{0,}[A-Z0-9!%{}^~_@#$♥]{1,}([.][A-Z0-9!%{}^~_@#$♥]{1,}){0,2}([\s]{1,}[A-Z0-9!%{}^~_@#$]{1,}){0,1}"
#End Region
#Region " IMAGE STRING DECALARATIONS "
    Public Const EyeString As String = "iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjAuMTZEaa/1AAACYElEQVQ4T6WUv24TQRDGLwYCNCAkQEFI2dm9vfg4yZUbS4i49QNgkBANjUusPIHfgIYoKYGKhqShQKJJmQoBAiEEDYI3ICilw/ft7ZzXxEFIRNrs7Ow3v5s/vssmk0kry7Kl4XB4Ktot2vTxrDb31E7jGMNz9GVL3W73DC/6/f5ptbnzTHs0GgVfai/SBqAK/xem/lDy32C9Xu98ae0a/sq+MedO0tJm6Sy5tUhQSLGeG3mBdejFHmHnOsxFdpxzN/+EcedqSlaB9/4CAE8QPCXEi7zxRjaxHudi39ZgO/XGPu06dzGFMb4ZCg9FUVyH+F3M5qe3Frr5nsF3G3cH4WFGPlbOrdLP9pDV9BA9uoxyPkXYUQrzxj1ApuOqqlbCGdBZG+xn+qnlCiUHkdhXDQxl0s+suPCgRwQ4I9/RkivMBr5QfgS/JiMAmeaa5HdnMMtSNuvM7AZ8L724sWZUiDxkIIBb6gtL8juhZF4i8J7Cwo4BMDNvzEZ4gPqxz4B2S2FBY+39kCH/hWCxeyrgNOkvy/Iazj80CP36xl7jblmHF/wi+4PB4Gzz6rEnlTErEHypAyHEUFg2fnNX65JlTBi1vFMY+8pJMymw5l+9drstEH6I0IM46WXV0I6wXxH29YZzBWG8J+PYq9fpdC6hp88QNK3BKE1kG23Yhu89lpb5nBlrXJMhS06dhNMurL2Vr8qOZkNQtHcxmPU0CY2b+9qkMBVy8eOQ5zl+fj6nTd8imPqZYfg4LIJxAGqr/yTYsVcvDVKB2v8CU20AKpTZaiCf1oiiHXoEjWqTuHg/af0Gh2dvChU5Q8wAAAAASUVORK5CYII="
    Public Const ClearTextString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjAuMjHxIGmVAAAB5ElEQVQ4T41TTUtVURRdzyLNSU1MCSppEChkkKbkrK9JOJBAhCai9u45R1DoD/T+QrOgQTl2IkiI75xzuzpIkSQcaCmlDSoJHEhBBEW91vnQR0+eumDDvXvttde+Z58L9E0cw5A+i4GsDkfFSNaEvL1K7Qkgb65B6DeQ5jkScyqWVMewbmHdIoTZRqJvAcregLQ/mPjNeHJAkxyEbYXUKzT7y+dPeGAuAclUPSd45JtI84sNnkIWz0RRGcq2sW4+iM0mJ7+NQqEmkO77k3SQ5E+Sf+jyAoPTDYEkkqyD4657sTRbdO9EoRTFuyhkx0mOkfzmmwg9jmHbCDFzhdO9Dc52HeplO6tzQVSJ0elaOrlJdsIkRjOW2aDE9zWotLs8djW4tSp9j8LvXuiD36xmW4BSFedKJOY6Hd+XG/BMRNYc2UMgTQ/jY3ReZXyOjV7786gOjifsHRZ+obtb1ZJ3lcWbbLLFcBt4BWUuRkEFRHqXBR+8m7ALyOvLPu8OTdle5jbiVLyFXON/kOl9El+D2CxgJL0QmYASp3PrE3b3XN5xWx2BTMx5ildJurXNQc2cC8Q+sIm79ntTPsPD+ZO8hZOnmXzMxHi42wcih3yxi6bux+v3l8/D/Q/JUn14OQKcsxcD/wAb8/HwyOWubQAAAABJRU5ErkJggg=="
    Public Const DropString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjAuMjHxIGmVAAAA3klEQVQ4T2MYBVCQvkeLIWNvLEP9fhaoCH4Qv18AyoKC9N0PGDL2/AfSyxlCVzFDRTEBSC5j9wyI2j11UFEgSN+zFCwIwQsZ6uuZoDIIELqKDah5NULd3ulQGSDI3cYOFNyKkNwzi4HhPyNUloGh8Bgninz67m0MaWe4oLJQEL+fA+iS3UiKpoDFk47wAtkH4OIZe1YCNbOC5TAAyNT0PUiK904GOvsEnJ++Zy7eMAKDpA28QMXHEIbANU9C8RZekLabH6jhFELz7laoDAkgcZsoUPMSIC6AigxPwMAAABk/eh6Y0kgKAAAAAElFTkSuQmCC"

    Public Const EditString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABl0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC4yMfEgaZUAAAFvSURBVDhPlZMrSARRFIbXRREfaFBUENuIsDM7D5YZ1uIUgwYtilksJvMiIgajoBg0iGnTBsFiEKNBRCzCVmERTMLCGlaDr+/oWVD0jvrDx73n3P+c+xgm9R85jtPved627/txLpdr0fTfFQTBFA2eoQYlmvTq0u9i12XXdVcpnIcT4gco63KyKIigDo9QgQLFkzScUItZHLObgnM4oKCPwnXmNcYjWVObWZjnKKyDIzFjhxwdjvP5fNu7yaQoirpocIV5TVPykNPEVYg1ZRamLShnMpkBTTXR8JJTbMr8I2UQhXi9O3Zc0lSauADXMKg5szDtw5llWa0SZ7PZEeJbWCFMS84ouR/Ge3Yf1ZQ03NDrdGrqZ8Vx3Iz5Ap4oKIZh2MNc7lOl4ZjakoX5lOJXgbkc+4b5Id+9XS3JomCn0UB54eUDXf4qFmZ4oNkGHNOCRYrkp6kwlmBc7d+FaRfD3idiGKbxgm3bQ8m/bSr1BlLvcgz+uCnlAAAAAElFTkSuQmCC"
    Public Const AddString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABl0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC4yMfEgaZUAAAF+SURBVDhPY8AFzMQ3lWqq1c3XUe6ebyFwTBcqTDwwFd20S0u9/r+ucvd/a8FTnlBh4oGp2OadMANsBU54Q4WJB2ZiW3bAXSBwyhcqTDwwE9u8HWJA138rgdP+UGHiAdAFW0EG6AANsOE/EwAVxgQWorvjzIR3zDQT3TnLXHTnbCCeayayc56R5LLHYANU2/6biW09YCa6fYmp6LZlQHo5EK8wF92x0kJkuxGDplrDfC31hv+YuJ4g1tCo9WbQVmmcr63W/B+MVVsQWK0JqrAByG8FuwSC2//rqECxars3gx3vVS8bnstltryXS+x4rxTZ8F4stOO9nG8muvkSyACQQhuBU3NseS6l2/NcTLXluZhix3shyZb3YoK90AkZaEhgAlPRLeshBnT8t+G7GAkVJh4AA24tzABbvkvRUGHigZnYtjVwF/BeiIUKEw/MRbevgoUB0L/xUGHigZnotpXwQOS9lAgVJh6Yi+1wBWbnTG2V1kw7nrOaUGE0wMAAABm0uZSULDJ3AAAAAElFTkSuQmCC"
    Public Const RemoveString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABl0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC4yMfEgaZUAAAHUSURBVDhPY4CBcsEz/EkM13mhXJwgg/uCWC7DNnYoFwJCGf4zn5SJXX1eNnJ3DMM/bqgwBkjnvCl9W8Hj0kbxqkaoEANDL98OodOysWv+acr//68p9/+afMDGmTwzRaDScLBavEP7gZLDDZCaXxoq/w/I5NTXM1xhY6jn3CV1X9HpFkgChP9oKP65KBu+F9kluZwXZR4qOd34qyn/D6bunGz4mniG+xxgBfX8h5UeKdmf+6upAFdwVS5gE8glyDaD8C8N5T9XFAI3ZzI8FARrhoGFIj1qjxQd7sAUglxyTjZsP7rNV+QDN7XzL0XVDAPYXALDIJuvAjXHM7wXgCrHDtBdAsN4bUYGID8/VHK8iW7ANXk/cJhAlWEHoNAGBRiyn2H4t4biX2DI7w1leMkDVY4KIDY7oYT2VYWALehhck3ef+NC3oXCUG0QAEph6DaDAgzkZ1CYPMCIHSSX1HNuk7mn6HoVpgAcz2ihDYkdBxSXnJOLXJnGcIaLoVT4CO9RmYSNMAlQIsEW2sgu+auh+H+HRAkwKe9nAUt6MPxjPy0bveSCXOh6jBSGBEAuuangfWSneHEDA8N/RqgwBHSL7+SuBzmJAKjnXy+QxjCTFcJjYAAACcw1wQVwmDAAAAAASUVORK5CYII="

    Public Const SortString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjAuMjHxIGmVAAABRklEQVQ4T+2TsUrDUBSGm9ykMUSJEGJAhDhkimDRLAUtwVcQUix0EJw6+AwZ3Vyd1cWH8AUsODi6CBYfQLqooNb/h3tqokPr7vAluec/5z8nyb2NsizNoiiUkCSJk+e59TMWhuFiEARLURR5Vb3BCwPEdd1VwzBOLMvaTdO0yVgcxwumaR4hfgEuwSny1sRkauA4zjrEa/AGhrZtt7Iss9kV63Mw0TxC26oZ6M4sftdJn+CWnbQBu1cNtqcGHBUjDiDcAxZK4gM4m2kgrwCTPsRXScT6kB/vLwZdiFWDHqf7N0ibvu8vY30lcfCEjdae2wDPsrnEgL96BJMd1s00wLibWN+BD609QzvmmahNoJTah/iik2jQpQG3M7p1ELsBYxbzu/x6Bc/zViAegD7v3N6MEz3JBprsyWmsGQjsJlTjghR9U6gvKYe7gZJsyxAAAAAASUVORK5CYII="
    Public Const CheckedString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjAuMjHxIGmVAAAB1klEQVQ4T2OgK/j37x/7nqMnDi/fubUIKkQauPb09vSZOlf+z1a78X/pro1VUGEIcHBwyHN0dIzChWfMnrGx2H7K/07xXf97RPf/79A/8H3pmaUiUO0MDPb29hFQJgb4//+/yMLJu95uYPrzfyPj3/8dErv+Rzm3F0KlIQCXAUDNzBt279u1iPfl/01AHsgAH8+kHQz/GRihSiAAlwEXHl3um6F1GawZhPuVTr0Gqg2FSiNAampq3snHZ+YAQ1oKKsTw7sdHv+lRh+Ca5wk/+pdaX+OHYcCbf2/4ctPrrkxVuvx/894Da0BiQKdzzO3d8QbkZJDmtSw//pfmTpscGhrKjGIAUCFTQ+Ginc0ym8EKZ0nc+RNfUxnR1Du7eAn3G7jtzUHrL4MMxTAABKo7J9dVaiyGK14g9vj/TOXrf2H8KUZnvs07tkQXpBarASCQm1m5ZoHIE7ghMLxA9Mn/aauXZUOV4TYAmFjiJk5ZuGaW0K3/y9negvESjhf/G6tm7nV3d5e0tbUFYycnJ2msBtjZ2akCU6N9y/TOBXXaa/5PE7jwvzS5946bm5sjSBwZ29jYiEK1YQJgQDHOWLWiqdF/3bcNV7cZQoVJB/ff31eAMvEABgYAPDoU9wQut5cAAAAASUVORK5CYII="
    Public Const CheckString As String = "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAYAAAAfSC3RAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABh0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMS40E0BoxAAAAN9JREFUOE+N0rsOAUEYhuF1SBwShUaicRN60aqIkkgU7kChUqJW0LkCpxAKd+BQaEm0DheglXg/WYK1MX/yZHcm+83OTH6LqmOBuYsZxvZzYs/1YI0wRRkFF0WU7Ocae1hDNBHUwKAG2OlFwQbcgiF0kYcHRsEAarhCAY2NgimcsEJCE9TPoB8xqDTe4oykJuxyBHWWDjZIow9tsQIt+CxHMIwcDrjgBi0UwXv93KoPWRyxRBTf5Xo5uvIM4o+Rsz6CLWirXgOvoFpOLaS2U0/+o2M8Wq4K/VUrmeBbq30HmqxP1SI+lSYAAAAASUVORK5CYII="
    Public Const UnCheckString As String = "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAYAAAAfSC3RAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAYdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjEuNBNAaMQAAABvSURBVDhPYwCCWiDeCcRbceAtQLweSm+Eis0FYoZ1UIEEII7CgaOBOBZKnwLim0DMsBaIW4CYA8QhAqwB4hsgxqhGTDBUNbYBMRcQMxGB4RpBSQ6UhDYBMShNEsJPgRic5IqBGGQryCRiMFAtw0QAbhE+zDCrvcQAAAAASUVORK5CYII="

    Public ReadOnly CheckImages As New Dictionary(Of String, String) From {
                         {"checkedBlack", "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAYAAAAfSC3RAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsEAAA7BAbiRa+0AAABwSURBVDhPlZABDoAgCEWh+x+0WxifxIHDgrcxBXxiMRENiTYqCljLMPO45r5NWcQUxExrogn+k37FTAJL3J8CThJYojXt8JcEwlOdfG+5hgeZ9OOtmOZrJklNV/TTn2NSNslANdxe4Tixgk58tx2IHtIlOgxG8FAIAAAAAElFTkSuQmCC"},
                         {"uncheckedBlack", "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAYAAAAfSC3RAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsEAAA7BAbiRa+0AAAAtSURBVDhPY2RgYPgPxCQDsEYgANFEA0ZGxv9MUDbJYFQjHjCqEQ8gM5EzMAAAoBMHFwfr1LQAAAAASUVORK5CYII="},
                         {"checkedWhite", "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAYAAAAfSC3RAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsEAAA7BAbiRa+0AAABnSURBVDhPxZJRDoAgDEM3739nXNEmbCko/vi+XOhLYdFbYB+g6Nf4mnbcH9ukRvdcXF7BAaHciCDDC6kjr/okgVFEIBmBlMAo8pDhqQTqVavcZyytLk69kQnZRORygmkT+e/P2cTsBCdlLwZDKAEtAAAAAElFTkSuQmCC"},
                         {"uncheckedWhite", "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAYAAAAfSC3RAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsEAAA7BAbiRa+0AAAAnSURBVDhPY/wPBAxkAJhGRgiXaPCfCcogGYxqxANGNeIBZCZyBgYAk5cNDhG2VLEAAAAASUVORK5CYII="}
                         }

    Public Const DataTableString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAmJJREFUOE99Ul1PGkEU3T60Puiv8iPR+OqrH0++m0ir0TQxJjWR4EdrCBgBG4wQSyEiQmhgXUAqKwtbwGRZ2Fl2MaAoghWILCxMh5AaNU3v070z98w55955AyHE/gbP85Ikoardbv9+qLx914NupXq9r7ev1YLtNuzt7cHQ0VOwLPuUcwCgHPUxLF+twrs7mM9Dkoy9ACSTyVcAVLIpvlbrAK6uYChEY7IMRbFB01W/v3R8nLNaBbM5ZTQm9HrSbA4TRIJNgVpNLpXkQkEOhaLY5aWcycjRaP30tOL1llyugt2es9nEvT1ma+tselqr1pgZBsTjgKYBQQSxTKbJcS3E4PPdOByi1crr9b+USu/8vGVy8vPg4AfDrv3hAd7ewlwOBoMUxvMNBAiHKy5X3mRKoYfVanJl5cfc3LeJic2Bgfe7Xx3lMry+htksDAQoTBAaglCnqNLRkWAwxLVaSqUilpYcCoV5aurL6OhH1ZoxFgORCDg/BzgewACQugxeb8FmQ47T29uR1dWOpC6DZvswn2+KYhOAps93jjw8ptPSPyWNj28MDy+oNYeiCDkOMgzEcRLjuEdBaFFU0e3OmkyMTkevr/uXl51dSUNDcxub39EO0YYuLqDHQyJJtf+YRlNaVeoIgsTxkMdzRlE0YqgJgkSSRYulox6ZVirxxUXbzMweYhgZWTw5oZ9/HzTWKsvKkUjF7y/a7Zf7+yyCra35FhasyHR/vyIaTb8AyHKbpks4fuN05p7vYXbWNDb2aWfH+by787W79f19PZG4JgjO4biwWKiDg59uN1UuV151o/IPPuNL2ItzNKQAAAAASUVORK5CYII="
    Public Const StarString As String = "iVBORw0KGgoAAAANSUhEUgAAAAwAAAAMCAIAAADZF8uwAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAXtJREFUKFNtkU1LAlEUhkcQbFt/xT/gylX+AJcRXDDCRS7CCCwCFzXgJsiLiYQwOWUZVx2/onGERO9cJ7OoRWbgJ+aIEgWGKF01AqGXszjvy3M4cI6m3W6rqqrVajUaDTOv8Xis0y0sLS0yGONyuVyv15vzarVatVoNYzmfx0wul6tWq43/RCdlWc5ms0wmk8HKg/z4Oqt3tacUe382mpQkKcOkUhIMlZyowuF+5K7T//x+fmleK90g6frkijd8k06nGf9FwmjjDDYuFH+jxGg0GgyGjUbnHD25OLLBomhUZLxngh5APfCbtkmCNIbDYfPjy+q7N1gIDfXAwSH0Cxnt4tZJ4Vaph3FVrrWOkmXgJEYg6gHLocQEMgK4yoqb3oLrtLR2rOwGlH1/EUBi3qMQ5NB0nWkHAihaIHFCAqdFG2oBKy7PoMBlxGxjgSO4DsOsG3k8AoTowI2oBfbgivUwkpImF+f5K54XeEEICUIsFhemDbU0RChOH/EDcKchcY4euAgAAAAASUVORK5CYII="

    Public Const BookClosed As String = "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAIAAACQKrqGAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAALDgAACw4BQL7hQQAAAddJREFUKFNtkd9P2lAcxYsPvvnv6njBYIAhEKVhxkHQoREFtOWHcYEZQKoIpSkt0dKglh+9GAwphlbBUSzsEhyyZOfp/vice3LPVzcej5E5jUZjRVF1OmRpaXFhQTd/hUB0puG7JtTkZArEYkKZl7rd3yNo/atPVNNGQlWOnwMsLJz9bMYTtdRFg2XbqqpN4Q8UcvWGEjmtneBVPAwIQk6nnn8l2jguRqL3HCdBwwSd5FZlyIUj9dBxPRGXotFW6Bj4/VWvl9/YYEzmPGgqyCS3Jp8nAXzvwC9gWDMYaPh8925PGUVLViv1ZSWztpZqtWSk03kLYRXf3u3Oj7sd7517m99ycc7NksPBWCyUXp8xmZNZku313xBRlLe+k55dat2RW7dR1q+UxUyajKRxNafXE1Zb+qZye3RSBEBGRPDs2SXyDB/AaLMFZuWMxvyqIQs5myPNcDeQc7kKktRD+v3B1fXDvr+QpfnLAv9tmzQYrlaWCTvkylwQKzqd5NPT60dZw6HWbr8eBpjwGctWhCOM3kQvqBIbwot2+/XLy+CfXuFGVd9zefEwyNDcA+g8YjEWRckZ9zmCqQ8W1xC7mUshgDFuDz3N/c9gZ0e93qD5qMB/zHNw/QdY3clc1dADtgAAAABJRU5ErkJggg=="
    Public Const BookOpen As String = "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAIAAACQKrqGAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAALDgAACw4BQL7hQQAAAeJJREFUKFN10V9v0lAYBnCWJV5544fxM3jntYlGs5joxbIlOhNHNEsVNS4zkexPtiUsm8ZZNlmxIRtsrJSuDih1qXQUgrUtpS3dpCtDkI5hwRYSuJg+d+c5vzfnTc5Au912dKJVGlcuXxocHOge/xGLWuGkX/t0KZuvNM7NbnMxDqtS1foSmNHKxhf6+IdctRrjrElli0GcRZJib8am6bQ25Sb1spHmNEY4zXA6tENHknm9Ykx7v9VqzfPOUzalqKOXr/d1vc5JJT/KfsZ4SdXOGk3r6tUi5t8QcOyoXm/alPyqAC7MogVFW92iITTXp/NBglTmZ9Nq8bdNiaQ8AaBdurSRWN9levTFXIA8KE6/TRWVmk3xGD/+NHTSoQsg5t2me/TZjM+iU5MHslR1tFqtIEKNjPkSRMGic6sRMJQSlZJyfBqn2GVfNE5Iz4GYKFYs2mZy+VGnB3CF0T125kN4PXwYJRj3SugdhB/mJAwXnOOoIJTtBf6YJisow48X3bN7bzybd50LgBvywgksnksxcmArO/pgk+NPbNoNX1DHJlau33JdvTZ047bn5tDynXvv7w+Dj5zwwydrvPizT03T/M7La3DkI7QD+ndBGPHCyKdANIiQMTJTrRl9+r+v7/V/AT5wyfCHirK9AAAAAElFTkSuQmCC"

    Public Const ArrowExpanded As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAM9JREFUOE/N0kkKg0AQBdCf+x/KlaALxXkeF04o4g0qXU1sDPQqBhLhI72o17+gH0SEWx8Dd3JrWLa/c/t3Ac/ziOM4Dtm2LWNZFpmmKWMYhsq1tVphmiYw0Pc9tW1LVVVRnueUpilFUURBEEjAdd23tdVhWRZckaZpqCxLiSRJIocFwpfogW3bwMg4juDqXdfRifCwQCCawPd9PbDvO9Z1VQgPMcJ/0QRZloGRMAz1wHEcOJF5njEMA14I6rpGURQSieNYD3z6Hv7oIf1shSf3G9UMQ+Vu/QAAAABJRU5ErkJggg=="
    Public Const ArrowCollapsed As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAANFJREFUOE+l00kKhDAQheHq+x/KlaALxXkeF04o4g1ep0J3Y0OEiILoIvnyK9QLAD26GHhyKzc7joNpmmhZFtq2jfZ9p+M4lGsvgTOyrqtEVKWXQN/3YGQcR1nCiDZg2zbatsUZmef5HlBVFZqmQdd1smQYBn3AsizkeY6yLP8Q7U8wTRNpmv4QLhAl+gUMRFGEJEnA76KE6rrWBwzDQBAE4KdAKMsyKoriHvBBSJTQF9H+B7zZdV3yPI9836cwDCmOY/2CO7PxaJDkJN85TbX2Db5d1YfJcQ3TAAAAAElFTkSuQmCC"

    Public Const LightOn As String = "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAIAAACQKrqGAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAAUtJREFUKFNj/P//PwMK+P/r6ytGRmZWLhFUcQYGoFI4+PryxMfbk388X/rj2ZIPt6d8e3MRWRah9MurC9+er/375fi/T3v/ftr99+vJr49Xfnt3C64aofTtrSV/Ph76/bju9/Npv59P+v245s+no29vLsWi9M3VGT/uVPy4lvTzQdfPe80/Lkf+vFvz+soMLEpfnu/4diHy683Wn6/W/3y96euNpm8Xol+c68ai9NnF+Z9Oer/cYfJ6n8Orfc6vdpp+Oun//PIyLEp///hyb1fS620Gd1ca3Ftl8Gab4f29eX9+/8SiFCj0/fPrh9tib6/2vLnU6N6mkF/fP2EPLIjoo2OTrq+JuDLf5NHRCcjqQDGFxr91fPn2/oC1zTa3Tm8koPTIkSN7du8qKS4+deoUAaWXL19uamoqKSm5ffs2AaVA6b9ggKYOyAUAkObu3QMxkwMAAAAASUVORK5CYII="
    Public Const LightOff As String = "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAIAAACQKrqGAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAAlwSFlzAAAOwwAADsMBx2+oZAAAARZJREFUKFOFkM1qg1AQhdOdT5uF6QM0O9MHaHYppFtXulJBgxupeiVWEfzb+IMUBK1YQRHtoKCmCJ7FZZj57rlzz1Pf97tHNU3TdR2GYf/6O0AnpWmqKIqu64ZhqKoahuFyOqN5ngNUVdXvoLIsPc+L43iiZxSciqIA459BUMBlTdNWULBMkiSKou9BUARBgBCCvUd6dgUD3/cdx4F3Qa7r2rYtSdKKq2VZ9/sdbOA0TVOWZag1hFbQtm1FURQEgaZpiqJ4nuc4Dn65gkKrrmuAPq7X89uZJMksy9bDGrvyp/x+uRAEwTDMknv41jj40vXjy3G/38MmGyhC6k24PR8OLMtuoJDX6+mE4zjEvIHCGDKfYl/Sf9M5/Uxpz2tBAAAAAElFTkSuQmCC"

    Public Const DefaultExpanded As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAO9JREFUOE+lk9sKAVEUhnkoD+E5vJJzuHEshzs1DiGUQ0RRigsKJceRcWZ+tppxmjXKTK2bqe/791p7Lz0AnaaPCbSUDI8mK3iiBbgjebjCOThCGdiDKVh9SZi9nFylWvue9wyVBZ5YEaIIXK4ijqcrtvsLeOGMGX/EeH7AYLJDdyjAYDQpC9yRwk+43d/QAnZstWQGN3prWuC890wdW4IrHZ4W2AJpuWfW52cxuNha0gKLP/E1sNdkBmebC1pg9nFv01aCU/W5ukC6KgrmqjN1AbtnNThentIC9sKUhvf5j3yJ/+6DpkV6bPK/yRJ3A/PE7e2oP8DgAAAAAElFTkSuQmCCAPjCzMoz/hO+xEPvwdYhbS75UGdNtwLNm+LI5h1FwAAAAABJRU5ErkJggg=="
    Public Const DefaultCollapsed As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAARpJREFUOE+lk9tKAmEUhfWhfIiew1eyUkkvPCaNdZdHbPCACIqCklZERQZCZhNZeZylS5jRhvlHcAb2zcD61tqLfzsBOGx9BNgZXdwffCKYLCEgFXF2IcN3XoA3nsNJJAtPOK1Ptd5Z+21NdUDwsgxVBRZLFdPZEj9/CyjjOd6VKd6GEzwPfnH/OobryG0OCEilvWK58SIGMLbRmW6ac+fpG1fynRjgX+9sjE0AY1PcfPiCVLAAnMby+s4UGqfWVZDI98QJjqMZ08LoTHG5PUI82xUDPJH0v7YZmyk08U3rA9HMHsBuYbvOFOcaQ4RTt9YJWFix2Ueq8rgpjDszNp0pDl1bAPjCzMoz/hO+xEPvwdYhbS75UGdNtwLNm+LI5h1FwAAAAABJRU5ErkJggg=="

    Public Const ChevronCollapsed As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAKCAYAAAC9vt6cAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjAuMjHxIGmVAAABRklEQVQoU2PABz5//uz06dMneyiXNHD8+PFyGxubf2ZmZn+PHj1aDhUmDgA1FxkbG/8HMsFYR0fn/5EjR4rAkoTAiRMnckxMTP4BmXADQBhoyL9jx47lANm4AVBzuqmpKVgzIyPjf3Z2djAGsUFienp6IEPSwYrRAVBzEsi/QOZ/Jiam/83NzQf+/fs3G4QnTJhwnJWVFWyIvr4+KEySwJpg4PXr17HJyclwzSUlJRv////PCpFlYADaytnU1LSbhYUFbEh0dPSfL1++xIIl7969K/cHCCZPnvxPQEDgX2lp6RagZjawJBI4c+YMF9CQ/YKCgv96enr+/f379+eVK1ckGG7evCkN5PwCavoPpHc/evSIE6oHA+zfv5/n+/fvR6Bqf9y5c0cMLPHx40err1+/ZuHTDANAW3mACSz7w4cPZgwMDAwA7Fq34WL8tRIAAAAASUVORK5CYII="
    Public Const ChevronExpanded As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAKCAYAAAC9vt6cAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABh0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMS42/U4J6AAAAUNJREFUKFNjAIGPHz9aff36NevRo0ecYAE84MqVKzyfP3/O/vDhgxlY4ObNm9J///799R8IgPRufIbs37+f5/v370egan/cuXNHjOHu3btyf4Bg8uTJ/wQEBP6VlpZuAcqzQfXAwZkzZ7iampr2CwoK/uvp6fkHNOAn0DUSYMnXr1/HJicn/wUy/zMxMf0vKSnZCDSEFSwJBMeOHeMEat7NwsLyH8j9Hx0d/efLly+xEFkoOHHiRJKZmRnckObm5gP//v2bDcITJkw4zsrKCtasr6//9+jRo0lgTegAaEi6qanpPyDzPyMj4392dnYwBrFBYnp6ev+ArkkHK8YFgIbkmJiYgA1Bxjo6OiDNOUA2YXD8+PEiY2NjZM3/jxw5UgSWJBYADSm3sbH5BwoXoJ/LocKkAWCCcfr06ZM9lIsFMDAAABo0t+GfVFaJAAAAAElFTkSuQmCC"

    Public Const PdfString As String = "iVBORw0KGgoAAAANSUhEUgAAABIAAAASCAYAAABWzo5XAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAYdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjEuNWRHWFIAAAH7SURBVDhPrZNBUhNBGIXnCIgX0BtwA72AVR7AQ7h2pQaIxAABN5ZEMkkqsbKOGzcWJRAVFdBsoyKo4BAksDGZzPh8r2s6zsSKWpZ/1VeT7pr+8vrvHgfAfyExqFarqFQqKJfLKJVKcN0iCoUC7i8vI5/P497SEorFEl9NSkRiIEkYhgMCEfxEf9RqtVBwXb7+G5GSSOD3+/D9Pnq+j17PR1d0e0bkeZ6RKWV8bUKk7UhUr9d/QbJms4lGo4FarYaFxTtcMkKknphENo2eNhH5xlSiHwTILSxyyZBo1TmDN+Q12SZbEZvkFXkZ8YJskOcRz8gjZ8wIjUiSjnMWx+QrOSJtcuiMwyNfyAHZJ5/JJ/KR7JEG1w5ESiKJLX9lHe2x89EI+N45wenVa0Ziq7uyht1hkbajJKrOxAWEO7s4uXzFjJWkzTnz+9yEeSqJJB/Ielykfmg7qq77wDy9KJHdTshUhxcvmTmlOc3dxQ7n1+I9kkg9UR3zZUmURCXJEdMFTKkkKiWR5D3XrcZFOh01VhVvrEoCf7uJNmXajspK3g2LdMR/Oh3bk7jkLXkSFz3mYNQ9eUp0MkKNVU+UQkjykAxEFl178539BTdSk0ZgSYh07edzOczOzSM7O4dMNouZzG2kZzKYTt/C5HQaqakpI7l+M8UlI0T/DpwfUyqMa1e21YsAAAAASUVORK5CYII="
    Public Const BlockString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAALEQAACxEBf2RfkQAAAB10RVh0Q29tbWVudABDcmVhdGVkIHdpdGggVGhlIEdJTVDvZCVuAAAAGHRFWHRTb3VyY2UASW50cmlndWUgSWNvbiBTZXSuJ6E/AAAAGHRFWHRTb2Z0d2FyZQBwYWludC5uZXQgNC4xLjb9TgnoAAAA10lEQVQ4T6WSsRHCMBAEPyQgcEgZFOCEiOJcCkU4IHQBFEBEwJgZChB3/L+wLInBJtiR/qSTTy9LCOEviuISclGkAR0YwGgj6ybbC1JBpAXXt5xzA22yX1eimV92cw94GLUDeADqPCRJEieYMialfqJtwdn0p41dXGc12cy7UtKYqfkCjjYfoodVLLRhlBh7bt6BjdVj7QBPwDunZl3fm1ZN4D242/gx6/rJ9GoPGJ1dpsyG8c6MzS+7ma9UeQWiT+eHzKH5y3/gaJKVf+IKiuISiuLvBHkBB+NzX3/RhhoAAAAASUVORK5CYII="
#End Region
    Public Enum Theme
        None
        White
        MediumBlue
        Blue
        DarkBlue
        MidnightBlue
        Pink
        Gray
        MediumGray
        DarkGray
        Black
        Red
        DarkRed
        Yellow
        Gold
        Green
        DarkGreen
        Orange
        Brown
        Turquoise
        Purple
    End Enum
    Friend Function ThemeToImage(colorTheme As Theme) As Image
        Return MyImages()("glossy" & colorTheme.ToString)
    End Function
    Public ReadOnly GlossyImages As New Dictionary(Of Theme, Image) From {
        {Theme.White, My.Resources.glossyWhite},
        {Theme.Black, My.Resources.glossyBlack},
        {Theme.Blue, My.Resources.glossyBlue},
        {Theme.MediumBlue, My.Resources.IBM},
        {Theme.DarkBlue, ShadeImage(My.Resources.glossyBlue, Color.Black, 128)},
        {Theme.MidnightBlue, ShadeImage(My.Resources.glossyBlue, Color.Black, 192)},
        {Theme.Brown, My.Resources.glossyBrown},
        {Theme.Green, My.Resources.glossyGreen},
        {Theme.DarkGreen, ShadeImage(My.Resources.glossyGreen, Color.Black, 128)},
        {Theme.Gray, My.Resources.glossyGrey},
        {Theme.MediumGray, ShadeImage(My.Resources.glossyGrey, Color.Black, 64)},
        {Theme.DarkGray, ShadeImage(My.Resources.glossyGrey, Color.Black, 128)},
        {Theme.Orange, My.Resources.glossyOrange},
        {Theme.Pink, My.Resources.glossyPink},
        {Theme.Purple, My.Resources.glossyPurple},
        {Theme.Red, My.Resources.glossyRed},
        {Theme.DarkRed, ShadeImage(My.Resources.glossyRed, Color.Black, 128)},
        {Theme.Turquoise, My.Resources.glossyTurquoise},
        {Theme.Yellow, My.Resources.glossyYellow},
        {Theme.Gold, ShadeImage(My.Resources.glossyYellow, Color.Goldenrod, 128)}
    }
    Friend Function GlossyForecolor(glossyTheme As Theme) As Color

        '/// .Net Gray colors are deceiving. Gray (128, 128, 128) is darker than DarkGray (169, 169, 169) White = (255, 255, 255)
        Select Case glossyTheme
            Case Theme.Gray '=.Net LightGray
                Return Color.Black
            Case Theme.MediumGray
                Return Color.White
            Case Theme.White
                Return Color.Black
            Case Else
                Dim glossyColor As Color = Color.FromName(glossyTheme.ToString)
                Return BackColorToForeColor(glossyColor)
        End Select
        Dim testColor As Color = Color.Gold

    End Function
    Public Function ShadeImage(imageIn As Image, OverlayColor As Color, Optional OverlayAlpha As Byte = 64) As Image

        If imageIn IsNot Nothing Then
            Using g As Graphics = Graphics.FromImage(imageIn)
                g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
                Using overlayBrush As New SolidBrush(Color.FromArgb(OverlayAlpha, OverlayColor))
                    g.FillRectangle(overlayBrush, New RectangleF(New Point(0, 0), imageIn.Size))
                End Using
            End Using
        End If
        Return imageIn

    End Function
    Public Function ResizeImage(image As Image, imageSize As Size) As Bitmap
        Return ResizeImage(image, imageSize.Width, imageSize.Height)
    End Function
    Public Function ResizeImage(image As Image, width As Integer, height As Integer) As Bitmap

        If image Is Nothing Then
            Return Nothing
        Else
            Dim destRect = New Rectangle(0, 0, width, height)
            Dim destImage = New Bitmap(width, height)
            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution)

            Using g = Graphics.FromImage(destImage)
                g.CompositingMode = Drawing2D.CompositingMode.SourceCopy
                g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
                g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
                g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
                Using wrapMode = New ImageAttributes()
                    wrapMode.SetWrapMode(Drawing2D.WrapMode.TileFlipXY)
                    g.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode)
                End Using
            End Using
            Return destImage
        End If

    End Function
    Public Function SameImage(Image1 As Image, Image2 As Image) As Boolean
        Return ImageToBase64(Image1, Imaging.ImageFormat.Bmp) = ImageToBase64(Image2, Imaging.ImageFormat.Bmp)
    End Function
    Public Function ImageToBase64(image As Image, Optional ImageFormat As Imaging.ImageFormat = Nothing) As String

        If image Is Nothing Then
            Return String.Empty
        Else
            If ImageFormat Is Nothing Then ImageFormat = ImageFormat.Bmp
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
    Public Function DrawProgress(fillValue As Object, Optional fillColor As Object = Nothing) As Image

        Dim wh As Integer = 150
        Dim bmp As New Bitmap(wh, wh)
        Dim defaultColor As Color = Color.Purple
        fillColor = If(fillColor, defaultColor)
        Dim fillBrushColor As Color = If(fillColor.GetType = GetType(Color), DirectCast(fillColor, Color), defaultColor)
        Dim value As Integer
        Dim isPercent As Boolean = False
        If fillValue?.GetType Is GetType(Double) Then
            value = CInt(DirectCast(fillValue, Double) * 100)
            isPercent = True
        ElseIf fillValue?.GetType Is GetType(Integer) Then
            value = DirectCast(fillValue, Integer)
        End If
        value = value Mod 101
        Using graphics As Graphics = Graphics.FromImage(bmp)
            With graphics
                .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                .FillRectangle(Brushes.Maroon, New Rectangle(0, 0, wh, wh))
                Dim max As Integer = 100
                Dim totalWidth = CInt(.VisibleClipBounds.Width)
                Dim totalHeight = CInt(.VisibleClipBounds.Height)
                Dim margin_all As Integer = 2
                Dim band_width = CInt(totalWidth * 0.1887)
                Dim workspaceWidth As Integer = totalWidth - (margin_all * 2)
                Dim workspaceHeight As Integer = totalHeight - (margin_all * 2)
                Dim workspaceSize = New Size(workspaceWidth, workspaceHeight)
                Dim upperLeftWorkspacePoint = New Point(margin_all, margin_all)
                Dim upperLeftInnerEllipsePoint = New Point(upperLeftWorkspacePoint.X + band_width, upperLeftWorkspacePoint.Y + band_width)
                Dim innerEllipseSize = New Size((CInt(totalWidth / 2) - upperLeftInnerEllipsePoint.X) * 2, (CInt(totalWidth / 2) - upperLeftInnerEllipsePoint.Y) * 2)
                Dim outerEllipseRectangle = New Rectangle(upperLeftWorkspacePoint, workspaceSize)
                Dim innerEllipseRectangle = New Rectangle(upperLeftInnerEllipsePoint, innerEllipseSize)
                Dim valueMaxRatio As Double = (value / max)
                Dim sweepAngle = CInt((valueMaxRatio * 360))
                Using progressFont As New Font("Calibri", If(isPercent, 24, 32), FontStyle.Bold)
                    Dim format As String = If(isPercent, FormatPercent(valueMaxRatio, 0), String.Format(InvariantCulture, "{0:00}", CInt(valueMaxRatio * 100)))
                    Dim measureString As SizeF = .MeasureString(format, progressFont)
                    Dim textPoint = New PointF(upperLeftInnerEllipsePoint.X + ((innerEllipseSize.Width - measureString.Width) / 2), upperLeftInnerEllipsePoint.Y + ((innerEllipseSize.Height - measureString.Height) / 2))
                    .Clear(Color.Transparent)
                    Using borderBrush As New SolidBrush(Color.Black)
                        Using borderPen As New Pen(borderBrush, 2)
                            Using fillBrush As New SolidBrush(fillBrushColor)
                                .DrawEllipse(borderPen, outerEllipseRectangle)
                                .FillPie(fillBrush, outerEllipseRectangle, 0, sweepAngle)
                                .FillEllipse(New SolidBrush(Color.GhostWhite), innerEllipseRectangle)
                                .DrawEllipse(borderPen, innerEllipseRectangle)
                                .DrawString(format, progressFont, Brushes.Black, textPoint)
                            End Using
                        End Using
                    End Using
                End Using
            End With
        End Using
        bmp.MakeTransparent(Color.Maroon)
        Return bmp

    End Function
    Public Function ExtensionToImage(path As String) As Image

        Dim kvp = GetFileNameExtension(path).Value
        Return If(kvp = ExtensionNames.PortableDocumentFormat, My.Resources.adobe,
                        If(kvp = ExtensionNames.Excel, My.Resources.Excel,
                        If(kvp = ExtensionNames.CommaSeparated, My.Resources.csv,
                        If(kvp = ExtensionNames.SQL, My.Resources.DDL,
                        If(kvp = ExtensionNames.Text, My.Resources.txt, My.Resources.Folder)))))

    End Function
    Public Function RotateImage(b As Bitmap, angle As Single) As Bitmap

        If b Is Nothing Then
            Return b
        Else
            'create a New empty bitmap to hold rotated image
            Dim returnBitmap As Bitmap = New Bitmap(b.Width, b.Height)
            'make a graphics object from the empty bitmap
            Using g As Graphics = Graphics.FromImage(returnBitmap)
                'move rotation point to center of image
                Dim dx As Single = CSng(b.Width / 2)
                Dim dy As Single = CSng(b.Height / 2)
                g.TranslateTransform(dx, dy)

                'rotate
                g.RotateTransform(angle)

                'move image back
                g.TranslateTransform(-dx, -dy)

                'draw passed in image onto graphics object
                g.DrawImage(b, New Point(0, 0))
            End Using
            Return returnBitmap
        End If

    End Function
#Region " RANDOM NUMBERS "
    Private ReadOnly Rnd As New Random()
    Public Function RandomNumber(Low As Integer, High As Integer) As Integer
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
    Public Function QuotientRound(Dividend As Long, Divisor As Long) As Long

        '/// Dividend ==> The dividend Is the number you are dividing up
        '/// Divisor  ==> The divisor Is the number you are dividing by
        '/// Quotient ==> The quotient Is the answer
        If Divisor = 0 Then
            Return 0
        Else
            Return Long.Parse(Split(CDec(Dividend / Divisor).ToString(InvariantCulture), ".")(0), InvariantCulture)
        End If

    End Function
    Public Function DoubleSplit(Number As Double) As KeyValuePair(Of Long, Double)

        Dim doubleString As String = Number.ToString(InvariantCulture)
        Dim doubleElements As String() = Split(doubleString, ".")

        If doubleElements.Count = 1 Then
            'Integer...Decimals come thru as 0.234
            Return New KeyValuePair(Of Long, Double)(Convert.ToInt64(Number), 0)
        Else
            Dim wholeNumber As Long = Convert.ToInt64(doubleElements.First, InvariantCulture)
            Dim partNumber As Double = Convert.ToDouble("." & doubleElements.Last, InvariantCulture) * If(Number < 0, -1, 1)
            Return New KeyValuePair(Of Long, Double)(wholeNumber, partNumber)
        End If

    End Function
    Public Function QuotientRemainder(Dividend As Long, Divisor As Long) As KeyValuePair(Of Long, Long)
        Dim qr = QuotientRound(Dividend, Divisor)
        '48 / 17 = 2, 14
        Return New KeyValuePair(Of Long, Long)(qr, Dividend - (qr * Divisor))
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
    Friend Function CursorDirection(Point1 As Point, Point2 As Point) As Cursor

        If Point1.X = Point2.X And Point1.Y = Point2.Y Then
            Return Cursors.Default

        ElseIf Point1.X = Point2.X And Point1.Y < Point2.Y Then
            Return Cursors.PanNorth

        ElseIf Point1.X = Point2.X And Point1.Y > Point2.Y Then
            Return Cursors.PanSouth

        ElseIf Point1.X < Point2.X And Point1.Y = Point2.Y Then
            Return Cursors.PanWest

        ElseIf Point1.X > Point2.X And Point1.Y = Point2.Y Then
            Return Cursors.PanEast

        ElseIf Point1.X < Point2.X And Point1.Y < Point2.Y Then
            Return Cursors.PanNW

        ElseIf Point1.X < Point2.X And Point1.Y > Point2.Y Then
            Return Cursors.PanSW

        ElseIf Point1.X > Point2.X And Point1.Y < Point2.Y Then
            Return Cursors.PanNE

        ElseIf Point1.X > Point2.X And Point1.Y > Point2.Y Then
            Return Cursors.PanSE

        Else
            Return Cursors.Default

        End If

    End Function
    Public Function OrderedMatch(inString As String, inList As List(Of String), Optional ignoreCase As Boolean = True) As List(Of String)

        If inString Is Nothing Or inList Is Nothing Then
            Return Nothing

        ElseIf inString.Any And inList.Any Then
            Dim listMatches As New List(Of String)
            inList.ForEach(Sub(item)
                               'inString=Agt
                               'item=August
                               If item IsNot Nothing Then
                                   Dim letterIndexes As New List(Of Integer)
                                   Dim addItem As String = item
                                   For s = 0 To inString.Length - 1
                                       Dim letterString As String = inString.Substring(s, 1)
                                       For l = 0 To item.Length - 1
                                           Dim letterList As String = item.Substring(l, 1)
                                           If String.Compare(letterString, letterList, ignoreCase, InvariantCulture) = 0 Then
                                               letterIndexes.Add(l)
                                               item = item.Remove(l, 1)
                                               item = item.Insert(l, BlackOut)
                                               Exit For
                                           End If
                                       Next
                                   Next
                                   If letterIndexes.Count = inString.Length Then
                                       'All input letters are found in the item, but also must be in the same order!
                                       Dim saveIndexes As New List(Of Integer)(letterIndexes)
                                       letterIndexes.Sort()
                                       If saveIndexes.SequenceEqual(letterIndexes) Then listMatches.Add(addItem)
                                   End If
                               End If
                           End Sub)
            Return listMatches
        Else
            Return Nothing
        End If

    End Function
    Public Function OrderedMonths(findMonth As String) As String

        If findMonth Is Nothing Then
            Return String.Empty

        ElseIf findMonth.Any Then
            Dim Months As New List(Of String)(Enumerable.Range(1, 12).Select(Function(m) MonthName(m)))
            Dim ignoreCase As Boolean = True
            Dim monthMatch As String = String.Empty
            Months.ForEach(Sub(item)
                               'findMonth=Agt
                               'item=August
                               If item IsNot Nothing Then
                                   Dim letterIndexes As New List(Of Integer)
                                   Dim addItem As String = item
                                   For s = 0 To findMonth.Length - 1
                                       Dim letterString As String = findMonth.Substring(s, 1)
                                       For l = 0 To item.Length - 1
                                           Dim letterList As String = item.Substring(l, 1)
                                           If String.Compare(letterString, letterList, ignoreCase, InvariantCulture) = 0 Then
                                               letterIndexes.Add(l)
                                               item = item.Remove(l, 1)
                                               item = item.Insert(l, BlackOut)
                                               Exit For
                                           End If
                                       Next
                                   Next
                                   If letterIndexes.Count = findMonth.Length Then
                                       'All input letters are found in the item, but also must be in the same order!
                                       Dim saveIndexes As New List(Of Integer)(letterIndexes)
                                       letterIndexes.Sort()
                                       If saveIndexes.SequenceEqual(letterIndexes) Then
                                           monthMatch = If(findMonth.Length = 3, addItem.Substring(0, 3), addItem)
                                           Exit Sub
                                       End If
                                   End If
                               End If
                           End Sub)
            Return monthMatch
        Else
            Return String.Empty
        End If

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
        If Timestamp Is Nothing Then
            Return Nothing
        Else
            If Regex.Match(Timestamp, "'20([0-9]{2}[-.]){6}[0-9]{6}'", RegexOptions.None).Success Then
                Dim Values As New List(Of Integer)(From v In Regex.Matches(Timestamp, "[0-9]{2,}", RegexOptions.IgnoreCase) Select Integer.Parse(DirectCast(v, Match).Value, InvariantCulture))
                Return New DateTime(Values(0), Values(1), Values(2), Values(3), Values(4), Values(5), Values(6), DateTimeKind.Local)
            Else
                Return Nothing
            End If
        End If

    End Function
    Public Function Db2DateToDate(Db2Date As String) As Date

        '2020-06-01
        Db2Date = Replace(Db2Date, "'", String.Empty)
        If Regex.Match(Db2Date, "20[0-9]{2}-[01][0-9]-[0123][0-9]", RegexOptions.None).Success Then
            Dim dateElements As String() = Split(Db2Date, "-")
            Dim Db2year As Integer = CInt(dateElements.First)
            Dim Db2Month As Integer = CInt(dateElements(1))
            Dim Db2Day As Integer = CInt(dateElements.Last)
            Return New Date(Db2year, Db2Month, Db2Day)
        Else
            Return Date.MinValue
        End If

    End Function
    Public Function DateToAccessString(DateValue As Date) As String

        '#4/1/2012#
        Dim Elements As New List(Of String) From {"#" + DateValue.Month.ToString(InvariantCulture),
            "/" + DateValue.Day.ToString(InvariantCulture),
            "/" + Format(DateValue.Year, "0000") + "#"}
        Return (Join(Elements.ToArray, String.Empty))

    End Function
    Public Function TimespanToString(ElapsedValue As TimeSpan) As String

        Dim Elements As New List(Of String) From {"'" + Format(ElapsedValue.Hours, "00"),
            ":" + Format(ElapsedValue.Minutes, "00"),
            ":" + Format(ElapsedValue.Seconds, "00"),
            "." + Format(ElapsedValue.Milliseconds, "0000000"),
            "'"}
        Return (Join(Elements.ToArray, String.Empty))

    End Function
    Public Function TimespanToString(timeSpans As List(Of TimeSpan)) As String

        If timeSpans Is Nothing Then
            Return Nothing
        Else
            Dim fixedTime As Date = Now
            Dim endTime As Date = fixedTime
            For Each timeSpan In timeSpans
                endTime = endTime.Add(timeSpan)
            Next
            Return TimespanToString(fixedTime, endTime)
        End If

    End Function
    Public Function TimespanToString(startTime As Date, endTime As Date) As String

        Dim period As TimeSpan = endTime.Subtract(startTime)
        With period
            Dim minutesSeconds As KeyValuePair(Of Long, Double) = DoubleSplit(.TotalMinutes)
            Dim minutes As Long = Math.Abs(minutesSeconds.Key)
            Dim seconds As Integer = Math.Abs(Convert.ToInt32(Math.Ceiling(60 * minutesSeconds.Value)))
            Select Case True
                Case minutes = 0 And seconds = 0
                    Return String.Empty

                Case minutes = 0
                    Return Join({seconds, "seconds"})

                Case minutes = 1
                    'Minutes is SINGULAR!
                    If seconds = 0 Then
                        Return Join({minutes, "minute"})
                    Else
                        Return Join({minutes, "minute"}) & ", " & Join({seconds, "seconds"})
                    End If
                Case Else
                    'Minutes is PLURAL!
                    If seconds = 0 Then
                        Return Join({minutes, "minutes"})
                    Else
                        Return Join({minutes, "minutes"}) & ", " & Join({seconds, "seconds"})
                    End If

            End Select
        End With

    End Function
    Public Function AverageSpan(Dates As List(Of Date)) As TimeSpan

        If Dates Is Nothing Then
            Return Nothing
        Else
            If Dates.Any Then
                If Dates.Count = 1 Then
                    Return Nothing
                Else
                    Try
                        Dim diff = Dates.Max.Subtract(Dates.Min)
                        Return TimeSpan.FromTicks(Convert.ToInt64(diff.Ticks / (Dates.Count - 1)))
                    Catch ex As InvalidOperationException
                        Return New TimeSpan
                    End Try
                End If
            Else
                Return Nothing
            End If
        End If

    End Function
    Public Function ApproximateEnd(RunningDates As List(Of Date), CollectionCount As Integer) As Date

        If RunningDates Is Nothing Then
            Return Nothing
        Else
            RunningDates.Sort()
            Dim averageTimespan As TimeSpan = AverageSpan(RunningDates)
            Dim firstDate As Date = RunningDates.First
            CollectionCount = {1, CollectionCount}.Max
            For d = {RunningDates.Count, CollectionCount}.Min To CollectionCount
                firstDate = firstDate.Add(averageTimespan)
            Next
            Return firstDate
        End If

    End Function
    Public Function LastDay(InDate As Date) As Date
        Return New Date(InDate.Year, InDate.Month, Date.DaysInMonth(InDate.Year, InDate.Month))
    End Function
    Public Function Db2ColumnNamingConvention(ColumnName As String) As String

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
    Public Function DB2TableNamingConvention(tableName As String) As String

        'https://www.sfu.ca/sasdoc/sashtml/accdb/z0455680.htm
        'A name can be from 1 to 18 characters long.
        'A name can start with a letter Or one of the following symbols: the dollar sign ($), the number (Or pound) sign (#), Or the at symbol (@).
        'A name can contain the letters A through Z, any valid letter with an accent (such as a), the digits 0 through 9, the underscore (_), the dollar sign ($), the number Or pound sign (#), Or the at symbol (@).
        'A name Is Not case-sensitive (for example, the table name CUSTOMERS Is the same as Customers), but object names are converted to uppercase when typed. If a name Is enclosed in quotes, then the name Is case-sensitive.
        'A name cannot be a DB2 Or an SQL reserved word, such as WHERE Or VIEW.
        'A name cannot be the same as another DB2 object that has the same type.
        'Schema And database names have similar conventions, except that they are Each limited To eight characters. For more information, see your DB2 SQL reference manual.

        If If(tableName, String.Empty).Length = 0 Then
            Return String.Empty

        Else
            Dim Letters = tableName.ToArray.Take(18) 'A name can be from 1 to 18 characters long.
            Dim LetterIndex As Integer = 0
            Dim NewLetters As New List(Of Char)
            For Each Letter As Char In Letters
                If LetterIndex = 0 Then
                    'FIRST CAN BE A-Z, $, #, @ - A name can start with a letter Or one of the following symbols: the dollar sign ($), the number (Or pound) sign (#), Or the at symbol (@)
                    If Regex.Match(Letter, "[^A-Z@$#]", RegexOptions.IgnoreCase).Success Then
                        NewLetters.Add(Chr(Asc(BlackOut)))
                    Else
                        NewLetters.Add(Letter)
                    End If

                Else
                    'SUBSEQUENT CAN BE A-Z,  0-9, _, $, #, @
                    If Regex.Match(Letter, "[^A-Z0-9_$#@]", RegexOptions.IgnoreCase).Success Then
                        NewLetters.Add(Chr(Asc(BlackOut)))
                    Else
                        NewLetters.Add(Letter)
                    End If

                End If
                LetterIndex += 1
            Next
            Dim reservedWords As New List(Of String)(Split(My.Resources.reservedWords, vbNewLine))
            Dim tableNameConvention As String = NewLetters.ToArray
            If reservedWords.Contains(tableNameConvention.ToUpperInvariant) Then
                Return StrDup(tableNameConvention.Length, BlackOut)
            Else
                Return tableNameConvention
            End If
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
    Public Function RegexMatches(InputString As String, Pattern As String, Optional Options As RegexOptions = RegexOptions.None, Optional ascendingIndex As Boolean = True) As List(Of Match)

        If InputString Is Nothing Or Pattern Is Nothing Then
            Return Nothing
        Else
            Dim matches As New List(Of Match)(From m In Regex.Matches(InputString, Pattern, Options) Select DirectCast(m, Match))
            If Not ascendingIndex Then
                matches.Sort(Function(y, x)
                                 Return x.Index.CompareTo(y.Index)
                             End Function)
            End If
            Return matches
        End If

    End Function
    Public Function RegexStringMatches(InputString As String, Pattern As String, Options As RegexOptions) As List(Of StringStartEnd)

        If InputString Is Nothing Or Pattern Is Nothing Then
            Return Nothing
        Else
            Return (From m In RegexMatches(InputString, Pattern, Options) Select New StringStartEnd(m)).ToList
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
    Public Function MeasureText(textIn As String, TextFont As Font, Optional adjustmentFactor As Double = 1.03) As Size

        If If(textIn, String.Empty).Any Or TextFont Is Nothing Then
            Dim gTextSize As SizeF
            Using g As Graphics = Graphics.FromImage(My.Resources.Plus)
                g.TextRenderingHint = Text.TextRenderingHint.AntiAlias
                Using sf As New StringFormat With {
                    .Trimming = StringTrimming.None
                }
                    gTextSize = g.MeasureString(textIn, TextFont, RectangleF.Empty.Size, sf)
                End Using
            End Using
            Return New Size(CInt(adjustmentFactor * gTextSize.Width), CInt(gTextSize.Height))

        Else
            Return New Size(0, 0)

        End If

    End Function
    Public Function MeasureText(textIn As String, textFont As Font, g As Graphics) As Size

        If If(textIn, String.Empty).Any Or textFont Is Nothing Then
            Dim characterRanges As CharacterRange() = {New CharacterRange(0, textIn.Length), New CharacterRange(0, 0)}
            Dim width As Single = 1000.0F
            Dim height As Single = 36.0F
            Dim layoutRect As RectangleF = New RectangleF(0.0F, 0.0F, width, height)
            Using sf As StringFormat = New StringFormat With {
                .FormatFlags = StringFormatFlags.NoWrap,
                .Alignment = StringAlignment.Near,
                .LineAlignment = StringAlignment.Center
                }
                sf.SetMeasurableCharacterRanges(characterRanges)
                g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
                Dim stringRegions As Region() = g.MeasureCharacterRanges(textIn, textFont, layoutRect, sf)
                Dim measureRect1 As RectangleF = stringRegions(0).GetBounds(g)
                Dim textTangle As Rectangle = Rectangle.Round(measureRect1)
                Return textTangle.Size
            End Using
        Else
            Return New Size(0, 0)
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
    Public Function CenterItem(parentLocation As Point, ItemSize As Size) As Point

        Dim ObjectCenter As New Size(Convert.ToInt32(ItemSize.Width / 2), Convert.ToInt32(ItemSize.Height / 2))
        parentLocation.Offset(-ObjectCenter.Width, -ObjectCenter.Height)
        Return parentLocation

    End Function
    Public Function DotProduct(arr1 As Integer(), arr2 As Integer()) As Integer

        If arr1 Is Nothing Or arr2 Is Nothing Then
            Return 0
        Else
            Return arr1.Zip(arr2, Function(d1, d2) d1 * d2).Sum()
        End If

    End Function
    Private Sub CrossProduct(vect_A As Integer(), vect_B As Integer(), cross_P As Integer())

        cross_P(0) = vect_A(1) * vect_B(2) - vect_A(2) * vect_B(1)
        cross_P(1) = vect_A(2) * vect_B(0) - vect_A(0) * vect_B(2)
        cross_P(2) = vect_A(0) * vect_B(1) - vect_A(1) * vect_B(0)

    End Sub
    Public Function InTriangle(checkPoint As Point, points As Point()) As Boolean

        If points Is Nothing Then
            Return Nothing
        Else
            Return InTriangle(checkPoint, points(0), points(1), points(2))
        End If

    End Function
    Public Function InTriangle(checkPoint As Point, points As PointF()) As Boolean

        If points Is Nothing Then
            Return Nothing
        Else
            Dim pointAx As Integer = Convert.ToInt32(points(0).X)
            Dim pointA As New Point(Convert.ToInt32(points(0).X), Convert.ToInt32(points(0).Y))
            Dim pointB As New Point(Convert.ToInt32(points(1).X), Convert.ToInt32(points(1).Y))
            Dim pointC As New Point(Convert.ToInt32(points(2).X), Convert.ToInt32(points(2).Y))
            Return InTriangle(checkPoint, pointA, pointB, pointC)
        End If

    End Function
    Public Function InTriangle(checkPoint As Point, pointA As Point, pointB As Point, pointC As Point) As Boolean

        Dim p As New Numerics.Vector2(checkPoint.X, checkPoint.Y)
        Dim p0 As New Numerics.Vector2(pointA.X, pointA.Y)
        Dim p1 As New Numerics.Vector2(pointB.X, pointB.Y)
        Dim p2 As New Numerics.Vector2(pointC.X, pointC.Y)

        If InLine(checkPoint, pointA, pointB) Or InLine(checkPoint, pointA, pointC) Or InLine(checkPoint, pointB, pointC) Then
            'Point is on the border edge
            Return True
        Else
            'Check if point is between edges
            Dim a = 0.5F * (-p1.Y * p2.X + p0.Y * (-p1.X + p2.X) + p0.X * (p1.Y - p2.Y) + p1.X * p2.Y)
            Dim sign = If(a < 0, -1, 1)
            Dim s = (p0.Y * p2.X - p0.X * p2.Y + (p2.Y - p0.Y) * p.X + (p0.X - p2.X) * p.Y) * sign
            Dim t = (p0.X * p1.Y - p0.Y * p1.X + (p0.Y - p1.Y) * p.X + (p1.X - p0.X) * p.Y) * sign
            Return s > 0 AndAlso t > 0 AndAlso (s + t) < 2 * a * sign
        End If

    End Function
    Public Function InLine(checkPoint As Point, pointA As Point, pointB As Point) As Boolean

        If checkPoint = pointA Or checkPoint = pointB Then
            Return True
        Else
            Dim xMin As Integer = {pointA.X, pointB.X}.Min
            Dim xMax As Integer = {pointA.X, pointB.X}.Max
            Dim yMin As Integer = {pointA.Y, pointB.Y}.Min
            Dim yMax As Integer = {pointA.Y, pointB.Y}.Max
            If pointA.X = pointB.X Then 'Vertical line
                Return checkPoint.X = pointA.X And checkPoint.Y >= yMin And checkPoint.Y <= yMax

            ElseIf pointA.Y = pointB.Y Then 'Horizontal line
                Return checkPoint.Y = pointA.Y And checkPoint.X >= xMin And checkPoint.X <= xMax

            Else
                Dim slope As Double = (pointA.Y - pointB.Y) / (pointA.X - pointB.X)
                Dim points As New List(Of Point)(From p In Enumerable.Range(yMin, yMax - yMin) Select New Point(p, CInt(p * slope)))
                Return points.Contains(checkPoint)

            End If
        End If

    End Function
    Public Function CamelFormat(words As String(), Optional CapFirst As Boolean = False) As String

        If words Is Nothing Then
            Return Nothing
        Else
            If words.Any Then
                Dim newWord As String = String.Empty
                Dim wordIndex As Integer = 0
                For Each word In words
                    If wordIndex = 0 Then
                        newWord &= StrConv(word, If(CapFirst, vbProperCase, vbLowerCase))
                    Else
                        newWord &= StrConv(word, vbProperCase)
                    End If
                    wordIndex += 1
                Next
                Return newWord
            Else
                Return String.Empty
            End If
        End If

    End Function
    Public Function CamelFormatSplit(word As String) As String()

        'Could be camelFormat or CamelFormat
        If word Is Nothing Then
            Return Nothing
        Else
            If word.Any Then
                Dim words As New List(Of String)
                Dim firstLetter As String = word.Substring(0, 1)
                words.AddRange(Regex.Split(word, "(?=[A-Z])", RegexOptions.None).Skip(If(firstLetter.ToUpperInvariant = firstLetter, 1, 0)))
                Return words.ToArray
            Else
                Return Nothing
            End If
        End If

    End Function

    Public Sub SetSafeControlPropertyValue(Item As Control, PropertyName As String, PropertyValue As Object)
        ThreadHelperClass.SetSafeControlPropertyValue(Item, PropertyName, PropertyValue)
    End Sub
    Public Sub SetSafeToolStripItemPropertyValue(Item As ToolStripItem, PropertyName As String, PropertyValue As Object)
        ThreadHelperClass.SetSafeToolStripItemPropertyValue(Item, PropertyName, PropertyValue)
    End Sub

    Private Enum CharacterType
        None
        Space
        NotSpace
    End Enum
    Public Function WordRectangles(stringIn As String, fontText As Font, Optional boundsImage As Rectangle = Nothing) As KeyValuePair(Of Size, SpecialDictionary(Of Integer, SpecialDictionary(Of Rectangle, String)))

        If If(stringIn, String.Empty).Any AndAlso fontText IsNot Nothing Then
#Region "■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ G E T   W O R D   S I Z E S"
            Dim wordList As New List(Of KeyValuePair(Of Size, String))
            If 0 = 0 Then
                RegexMatches(stringIn, "[^ ]{1,}", RegexOptions.None).ForEach(Sub(word)
                                                                                  wordList.Add(New KeyValuePair(Of Size, String)(MeasureText(word.Value, fontText), word.Value))
                                                                              End Sub)
            Else
                Dim firstLetter As Char = stringIn.First
                Dim lastType As CharacterType = If(TrimReturn(firstLetter).Any, CharacterType.NotSpace, CharacterType.Space)
                Dim typeString As String = String.Empty
                For Each letter As Char In stringIn
                    Dim currentType As CharacterType = If(TrimReturn(letter).Any, CharacterType.NotSpace, CharacterType.Space)
                    If lastType <> currentType Then
                        wordList.Add(New KeyValuePair(Of Size, String)(MeasureText(typeString, fontText), typeString))
                        lastType = currentType
                        typeString = String.Empty
                    End If
                    typeString &= letter
                Next
                wordList.Add(New KeyValuePair(Of Size, String)(MeasureText(typeString, fontText), typeString)) 'Very last character of stringIn
            End If
            '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ M A X   V A L U E S
            Dim widthSpace As Integer = MeasureText(" ".ToUpperInvariant, fontText).Width
            Dim heightRow As Integer = wordList.Max(Function(w) w.Key.Height)
            Dim widthLargestWord As Integer = wordList.Max(Function(w) w.Key.Width)
#End Region
#Region "■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ G E T   W I D T H  +  H E I G H T "
            '■■■■■■■■■■■■■■■■ A long word may require an adjustment to the widthProposed depending on what line it shows. If next to an image, the width may need expanding
            Const width2height_Ratio As Byte = 3 'Looks best when width is 3 times the height
            Dim sizeText As Size = MeasureText(stringIn, fontText)
            Dim areaText As Integer = sizeText.Width * sizeText.Height
            '■■■■■■■■■■■■■■■■ Math explaining how height = Math.Sqrt(TextArea / width2height_Ratio)
            'Area = Width * Height                       ex Area = 10,000 ( x * y )
            '∵ Width = 3*Height                         ex x = 3y
            'Area = Width (3*Height) * Height            ex Area = 3y * y
            'Area = 3*Height * Height                    ex y² * 3 = 10,000
            'Area = Height² * 3                          ex y² = 10,000 / 3
            'Height = √Area/3                            ex y = √3,333.33   57.73
            'Width = Height * 3                          ex x = 57.73 * 3 = 173.21   ... 57.73 * 173.21 = 10,000
            'Area of 10,000 should have a width of 173.21 and a height of 57.73
            Dim heightProposed As Double = Math.Sqrt(areaText / width2height_Ratio)
            Dim widthProposed As Integer = CInt(Math.Ceiling(heightProposed * width2height_Ratio))
            If widthLargestWord > widthProposed Then
                widthProposed = CInt(10 * Math.Ceiling(widthLargestWord / 10))
                heightProposed = CInt(10 * Math.Ceiling(widthProposed / width2height_Ratio / 10))
            End If
#End Region
            Const padV As Integer = 5
            Const padH As Integer = 5
            Const padLine As Integer = 2
            Const consoleWrite As Boolean = False
            Dim rows As New SpecialDictionary(Of Integer, SpecialDictionary(Of Rectangle, String))
            Dim sizeWords As New Size(widthProposed, CInt(heightProposed))
            Const maxAttempts As Integer = 5
            Dim dataConsole As String = String.Empty

            For countAttempts = 1 To maxAttempts
                Dim attemptFailed As Boolean = False
                Dim indexWord As Integer = 0
                Dim boundsRollingLeft As Integer = 0
                Dim boundsRollingTop As Integer = 0
                Dim maxWidth As Integer = 0
                Dim maxHeight As Integer = 0
                Dim indexLine As Integer = 0
                Dim indexOfWordInLine As Integer = 0

                If consoleWrite Then Console.WriteLine($"{StrDup(25, "»")} Attempt#{countAttempts.ToString("00", InvariantCulture)}, Size{sizeWords}")
                wordList.ForEach(Sub(word)
                                     Dim sizeWord As Size = word.Key
                                     Dim stringWord As String = word.Value
                                     Dim sizeWordNext As Size = If(indexWord + 1 < wordList.Count, wordList(indexWord + 1).Key, New Size)
                                     Dim stringWordNext As String = If(indexWord + 1 < wordList.Count, wordList(indexWord + 1).Value, Nothing)
                                     Dim isReturn As Boolean = {vbNewLine, vbCrLf, vbCr}.Contains(stringWord)
                                     Dim isSpace As Boolean = Not (isReturn Or Trim(stringWord).Any)
                                     Dim isImageLine As Boolean = boundsImage.Height <> 0 And boundsImage.Width <> 0 AndAlso boundsRollingTop < boundsImage.Bottom
                                     Dim textStartsAt As Integer = padH + If(isImageLine, boundsImage.Right, 0)

                                     If Not rows.ContainsKey(indexLine) Then rows.Add(indexLine, New SpecialDictionary(Of Rectangle, String))
                                     boundsRollingLeft = If(indexOfWordInLine = 0, textStartsAt, boundsRollingLeft)
                                     boundsRollingTop = padV + indexLine * (heightRow + padLine)

                                     Dim widthAvailable As Integer = widthProposed - boundsRollingLeft
                                     Dim thisWordFits As Boolean = widthAvailable >= sizeWord.Width
                                     Dim nextWordFits As Boolean = widthAvailable >= sizeWord.Width + sizeWordNext.Width
                                     Dim isLastWord As Boolean = Not nextWordFits

                                     If consoleWrite Then
                                         dataConsole = $"{StrDup(5, "»")} Word:{stringWord}, Fits={If(thisWordFits, "Y", "N")}, Next:{stringWordNext}, Fits={If(nextWordFits, "Y", "N")}"
                                         Console.WriteLine(dataConsole)
                                     End If

                                     If isSpace And (indexOfWordInLine = 0 Or isLastWord) Then
                                         'D O N ' T   A D D   S P A C E S   A T   L I N E   S T A R T   O R   L I N E   E N D !
                                         boundsRollingLeft -= widthSpace
                                     Else
                                         Dim boundsWord As New Rectangle(boundsRollingLeft, boundsRollingTop, sizeWord.Width, heightRow)
                                         rows(indexLine).Add(boundsWord, stringWord)
                                         If Not thisWordFits Then
                                             '■■■■■■■■■■■■■ Is an issue since loop looks ahead one word to see if it fits before starting a new line
                                             'Ends up here when a long word is in the image zone and the combined width exceeds the proposed width
                                             'Leave on this row and try again with a wider value
                                             '■■■■■■■■■■■■■ Widen the widthProposed
                                             attemptFailed = True
                                             widthProposed = CInt(10 * Math.Ceiling((boundsRollingLeft + sizeWord.Width) / 10)) 'Rounded up to a 10 value
                                         End If
                                         If consoleWrite Then
                                             dataConsole = $"Index:{indexWord.ToString("00", InvariantCulture)}, Row:{indexLine.ToString("00", InvariantCulture)}, #inRow:{indexOfWordInLine.ToString("00", InvariantCulture)}, (Left, Width, Right, Bottom)({boundsWord.Left.ToString("000", InvariantCulture)}, {boundsWord.Width.ToString("000", InvariantCulture)}, {boundsWord.Right.ToString("000", InvariantCulture)}, {boundsWord.Bottom.ToString("000", InvariantCulture)})"
                                             Console.WriteLine(dataConsole)
                                         End If
                                     End If
                                     If isReturn Or Not nextWordFits Then
                                         '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ N E W   L I N E
                                         indexOfWordInLine = 0
                                         indexLine += 1
                                     Else
                                         boundsRollingLeft += sizeWord.Width
                                         indexOfWordInLine += 1
                                     End If
                                     indexWord += 1
                                     If consoleWrite Then Console.WriteLine(StrDup(50, "_"))
                                 End Sub)
                If Not attemptFailed Or countAttempts = maxAttempts Then
                    Exit For
                Else
                    rows.Clear()
                End If
            Next
#Region "■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ CALCULATE BEST FIT - NOTHING IS BEING DONE WITH THIS!!! "
            If 0 = 1 Then
                Dim pastImage = True
                If pastImage Then
#Region "■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ LINES THAT GO PAST THE ICON "
                    Dim iconAdjust As Integer = 0
                    Dim textAdjust As Integer = 0
                    '/// Key = Icon.Height delta, Value=Text.Height delta ... both must grow only as shrinking the Icon or Text height not ideal
                    '/// 4 possible outcomes: a) Neither change, b) Text grows, c) Icon grows or d) both grow
                    Dim iconHeight As Integer = boundsImage.Height

                    Dim qr = QuotientRemainder(iconHeight, heightRow) 'renders ==> (#Rows of text, #Pixels total remaining)
                    Dim countRows As Byte = CByte(qr.Key)
                    Dim pixels As Byte = CByte(qr.Value)
                    '(48, 17)=(2, 14) meaning 2 rows with 14 pixels to split between the 2 rows ( 7 each - too high ) ... additional row is just past the Icon bottom
                    '(48, 23)=(2, 2)  meaning 2 rows with 2 pixels to split between the 2 rows ( 1 each - OK ) ... text line height is just short of the icon bottom
                    '/// ∴ Low remainder = grow Text while high remainder = grow Icon

                    Dim textPixelsGrow = QuotientRemainder(pixels, countRows) 'Determines how to distribute pixels...(#Pixels, #Rows) ==> (14 pixels, 2 rows) ==> ( 7 pixels, 0 remainder)

                    If textPixelsGrow.Key <= 4 Then
                        'OK to use a hard value of 4 since padding 2 above text and 2 below text is ok, more than that is noticeable
                        iconAdjust = Convert.ToInt32(textPixelsGrow.Value)
                        textAdjust = Convert.ToInt32(textPixelsGrow.Key)

                    Else
                        Dim iconDelta As Integer = heightRow - pixels 'If textHeight=17 and pixels=14 then only 3 change
                        'Try evenly splitting pixels among the Icon and Rows
                        Dim pixelSplit = QuotientRemainder(iconDelta, countRows + 1) '...say 4 delta amoung 2 rows and Icon
                        Dim textGrowMax As Long = {pixelSplit.Key, 4}.Max
                        Dim iconGrowValue As Long = textGrowMax - pixelSplit.Key
                        iconAdjust = Convert.ToInt32(iconGrowValue)
                        textAdjust = Convert.ToInt32(textGrowMax)

                    End If
#End Region
                Else
#Region "■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ LINES THAT ARE BELOW THE ICON "
                    Dim textHeight As Integer = rows.Count * heightRow
                    Dim extraSpace As Integer = CInt(QuotientRound(boundsImage.Height - textHeight, 2))
                    If extraSpace > 0 Then
                        Dim newBounds As New Dictionary(Of Rectangle, String)
                        For Each row In rows
                            With row.Value
                                'newBounds.Add(New Rectangle(.Left, .Top + extraSpace, .Width, .Height), textBound.Value)
                            End With
                        Next
                        'TextBounds.AddRange(newBounds)
                    End If
#End Region
                End If
            End If
#End Region
            If rows.Any Then
                Dim widthMax As Integer = rows.Max(Function(row) row.Value.Max(Function(r) r.Key.Right))
                Dim heightMax As Integer = rows.Max(Function(row) row.Value.Max(Function(r) r.Key.Bottom))
                Dim widthRound As Integer = widthMax + padH
                Dim heightRound As Integer = heightMax + padV
                sizeWords = New Size(widthRound, heightRound)
            End If
            Return New KeyValuePair(Of Size, SpecialDictionary(Of Integer, SpecialDictionary(Of Rectangle, String)))(sizeWords, rows)
        Else
            Return Nothing
        End If

    End Function
    Public Function DisplayScale() As Single

        Dim VERTRES As Integer = 10
        Dim DESKTOPVERTRES As Integer = 117

        Dim g As Graphics = Graphics.FromHwnd(IntPtr.Zero)
        Dim desktop As IntPtr = g.GetHdc()
        Dim LogicalScreenHeight As Integer = NativeMethods.GetDeviceCaps(desktop, VERTRES)
        Dim PhysicalScreenHeight As Integer = NativeMethods.GetDeviceCaps(desktop, DESKTOPVERTRES)
        Dim ScreenScalingFactor As Single = CSng(PhysicalScreenHeight / LogicalScreenHeight)
        Return ScreenScalingFactor

    End Function

#Region " COLOR "
    Function ColorToHtmlHex(color As Color) As String
        Return String.Format(InvariantCulture, "#{0:X2}{1:X2}{2:X2}", color.R, color.G, color.B)
    End Function
    Public Function HtmlToColor(colorString As String) As Color

        If colorString Is Nothing Then
            Return Nothing
        Else
            Dim htmlColorRegex As Regex = New Regex("^#((?'R'[0-9a-f]{2})(?'G'[0-9a-f]{2})(?'B'[0-9a-f]{2}))" & "|((?'R'[0-9a-f])(?'G'[0-9a-f])(?'B'[0-9a-f]))$", RegexOptions.Compiled Or RegexOptions.IgnoreCase)
            Dim match = htmlColorRegex.Match(colorString)
            If match.Success Then
                Return Color.FromArgb(ColorComponentToValue(match.Groups("R").Value), ColorComponentToValue(match.Groups("G").Value), ColorComponentToValue(match.Groups("B").Value))
            Else
                Return Nothing
            End If
        End If

    End Function
    Private Function ColorComponentToValue(component As String) As Integer

        Debug.Assert(component IsNot Nothing)
        Debug.Assert(component.Length > 0)
        Debug.Assert(component.Length <= 2)
        If component.Length = 1 Then
            component += component
        End If
        Return Integer.Parse(component, NumberStyles.HexNumber, InvariantCulture)

    End Function
    Public Function BackColorToForeColor(backColor As Color) As Color

        With backColor
            Dim rFactor As Double = Math.Pow(.R, 2) * 0.241
            Dim gFactor As Double = Math.Pow(.G, 2) * 0.691
            Dim bFactor As Double = Math.Pow(.B, 2) * 0.068
            Dim colorThreshold = Math.Sqrt(rFactor + gFactor + bFactor)
            If colorThreshold < 130 Then
                Return Color.White
            Else
                Return Color.Black
            End If
        End With

    End Function
    Public Function FontImages() As Dictionary(Of Font, Image)

        Dim FontImageCollection As New Dictionary(Of Font, Image)
        REM /// INITIALIZE THEM
        Using ifc As Text.InstalledFontCollection = New Text.InstalledFontCollection()
            For Each availableFont In ifc.Families
                Using bmpFont As New Font(availableFont, 9, FontStyle.Regular)
                    Dim fontSize As Size = MeasureText(availableFont.Name, bmpFont)
                    Dim _Image As New Bitmap(fontSize.Width, fontSize.Height)
                    Using G As Graphics = Graphics.FromImage(_Image)
                        G.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                        G.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                        G.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
                        G.TextRenderingHint = Text.TextRenderingHint.SingleBitPerPixelGridFit
                        Using Brush As New SolidBrush(Color.White)
                            G.DrawRectangle(Pens.Black, 0, 0, _Image.Width - 1, _Image.Height - 1)
                            G.FillRectangle(Brush, 2, 2, _Image.Width - 4, _Image.Height - 4)
                        End Using
                        Using Format As New StringFormat With {
        .Alignment = StringAlignment.Center,
        .LineAlignment = StringAlignment.Center
    }
                            G.DrawString("A", bmpFont, Brushes.Black, New Rectangle(New Point(0, 0), fontSize), Format) 'availableFont.Name
                        End Using
                    End Using
                    FontImageCollection.Add(bmpFont, _Image)
                End Using
            Next
        End Using
        Return FontImageCollection

    End Function
    Public Function ColorImages() As Dictionary(Of Color, Image)

        Dim ColorImageCollection As New Dictionary(Of Color, Image)
        REM /// INITIALIZE THEM
        For Each colorName In ColorNames()
            Dim _Image As New Bitmap(16, 16)
            Dim colorValue As Color = Color.FromName(colorName)
            Using G As Graphics = Graphics.FromImage(_Image)
                Using Brush As New SolidBrush(colorValue)
                    G.DrawRectangle(Pens.Black, 0, 0, _Image.Width - 1, _Image.Height - 1)
                    G.FillRectangle(Brush, 2, 2, _Image.Width - 4, _Image.Height - 4)
                End Using
            End Using
            ColorImageCollection.Add(colorValue, _Image)
        Next
        Return ColorImageCollection

    End Function
    Public Function ColorNames() As List(Of String)

        Dim ColorType As Type = Color.Beige.GetType
        Dim ColorList() As PropertyInfo = ColorType.GetProperties(BindingFlags.Static Or BindingFlags.DeclaredOnly Or BindingFlags.Public)
        Return New List(Of String)(From CL In ColorList Select CL.Name)

    End Function
    Public Function ChangeImageColor(bmp As Bitmap, OldColor As Color, NewColor As Color) As Image

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
#End Region

    Public Function CookieToDictionary(stringValue As String) As SpecialDictionary(Of String, String)

        If stringValue Is Nothing Then
            Return Nothing
        Else
            Dim namesValues As New SpecialDictionary(Of String, String)
            For Each nameValue In Split(stringValue, ";")
                Dim kvp As String() = Split(nameValue, "=")
                namesValues.Add(Trim(kvp.First), Trim(kvp.Last))
            Next
            Return namesValues
        End If

    End Function
    Public Function GetPreferredBrowser() As KeyValuePair(Of String, String)
        Const userChoice As String = "Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice"
        Dim nameAndPath As KeyValuePair(Of String, String) = New KeyValuePair(Of String, String)("Unknown", String.Empty)

        Using userChoiceKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(userChoice)

            If userChoiceKey IsNot Nothing Then
                Dim progIdValue As Object = userChoiceKey.GetValue("Progid")

                If progIdValue IsNot Nothing Then
                    Dim progId As String = progIdValue.ToString()

                    If progId = "ChromeHTML" Then
                        nameAndPath = New KeyValuePair(Of String, String)("Chrome", "C:\Program Files\Google\Chrome\Application\Chrome.exe")
                    ElseIf progId = "FirefoxURL" Then
                        nameAndPath = New KeyValuePair(Of String, String)("Firefox", "C:\Program Files\Mozilla Firefox\firefox.exe")
                    ElseIf progId = "IE.HTTP" Then
                        nameAndPath = New KeyValuePair(Of String, String)("InternetExplorer", "C:\Program Files\Internet Explorer\iexplore.exe")
                    ElseIf progId = "AppXq0fevzme2pys62n3e0fbqa7peapykr8v" Then
                        nameAndPath = New KeyValuePair(Of String, String)("Edge", "C:\Program Files (x86)\Microsoft\Edge\msedge.exe")
                    ElseIf progId = "OperaStable" Then
                        nameAndPath = New KeyValuePair(Of String, String)("Opera", Nothing)
                    ElseIf progId = "SafariHTML" Then
                        nameAndPath = New KeyValuePair(Of String, String)("Safari", Nothing)
                    End If
                End If
            End If
        End Using

        Return nameAndPath
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
    Public Function LevenshteinDistance(s As String, t As String, Optional caseInsensitive As Boolean = False) As Integer

        s = If(s, String.Empty)
        s = If(caseInsensitive, s.ToUpperInvariant, s)
        t = If(t, String.Empty)
        t = If(caseInsensitive, t.ToUpperInvariant, t)
        Dim n As Integer = s.Length
        Dim m As Integer = t.Length
        Dim d As Integer()() = New Integer(n)() {} 'New Integer(n, m) {}

        If n = 0 Then Return m
        If m = 0 Then Return n

        Dim i As Integer = 0
        While i <= n
            d(i) = New Integer(m) {}
            d(i)(0) = Math.Min(Threading.Interlocked.Increment(i), i - 1)
        End While

        Dim j As Integer = 0
        While j <= m
            d(0)(j) = Math.Min(Threading.Interlocked.Increment(j), j - 1)
        End While

        For x As Integer = 1 To n
            For y As Integer = 1 To m
                Dim cost As Integer = If((t(y - 1) = s(x - 1)), 0, 1)
                d(x)(y) = Math.Min(Math.Min(d(x - 1)(y) + 1, d(x)(y - 1) + 1), d(x - 1)(y - 1) + cost)
            Next
        Next

        Return d(n)(m)

    End Function
    Public Function ConsecutiveLetters(word1 As String, word2 As String) As KeyValuePair(Of Integer, String)

        word1 = If(word1, String.Empty)
        word2 = If(word2, String.Empty)

        Dim shortestWord As String = If(word1.Length < word2.Length, word1, word2)
        Dim longestWord As String = If(word1.Length > word2.Length, word1, word2)

        Dim letterCount As Integer = 0
        For i = 0 To shortestWord.Length - 1
            If shortestWord.Substring(i, 1) = longestWord.Substring(i, 1) Then
                letterCount += 1
            Else
                Exit For
            End If
        Next
        Return New KeyValuePair(Of Integer, String)(letterCount, shortestWord.Substring(0, letterCount))

    End Function
    Public Function Anacronym(words As String) As String

        If words Is Nothing Then
            Return Nothing

        ElseIf words.Any Then
            Dim wordList As New List(Of String)(From w In Regex.Split(words, "[\s]{1,}") Where w.Any)
            Dim startingLetters As New List(Of String)(From w In wordList Where Not {"AND", "OF"}.Contains(w) Select w.First + String.Empty)
            Return Join(startingLetters.ToArray, String.Empty)
        Else
            Return String.Empty
        End If

    End Function
    Public Function AbbreviationCompare(word1 As String, word2 As String) As Integer

        word1 = If(word1, String.Empty).ToUpperInvariant
        word2 = If(word2, String.Empty).ToUpperInvariant

        '/// COMPARE LEFT AGAINST RIGHT 1st ///
        Dim abbreviationLeft As Match = Regex.Match(word1, "[^.\s]{1,}(?=\.)", RegexOptions.None)
        If abbreviationLeft.Success Then
            Dim matchAbbreviation As Match = Regex.Match(word2, "(?<=^|\s)CH[A-Z]{1,}")
            '/// Don't need to test on matchAbbreviation.Success since String.Remove(0, 0) does not throw an error ///
            Dim leftWord As String = Regex.Replace(word1, "[^.\s]{1,}\.", String.Empty) 'Removing abbreviation including .
            Dim rightWord As String = word2.Remove(matchAbbreviation.Index, matchAbbreviation.Length) 'Remove only 1 instance
            Dim theLeftMatches = RegexMatches(word1, "(?<=^|\s)THE", RegexOptions.None)
            Dim theRightMatches = RegexMatches(word2, "(?<=^|\s)THE", RegexOptions.None)
            If Not theLeftMatches.Count = theRightMatches.Count Then
                'Don't let a fluff word like <The> muddy the correlation
                For Each leftMatch In theLeftMatches
                    leftWord = leftWord.Remove(leftMatch.Index, leftMatch.Length)
                Next
                For Each rightMatch In theRightMatches
                    rightWord = rightWord.Remove(rightMatch.Index, rightMatch.Length)
                Next
            End If
            If leftWord = rightWord Then
                Return leftWord.Length
            Else
                leftWord = Regex.Replace(leftWord, "[^A-Z]", String.Empty, RegexOptions.None)
                rightWord = Regex.Replace(rightWord, "[^A-Z]", String.Empty, RegexOptions.None)
                If leftWord = rightWord Then
                    Return leftWord.Length
                Else
                    Dim shortestLongestWord = ShortestLongest(leftWord, rightWord)
                    Dim shortestWord As String = shortestLongestWord.First
                    Dim longestWord As String = shortestLongestWord.Last
                    Dim wordDeviation As Integer = LevenshteinDistance(leftWord, rightWord)
                    If shortestWord.Length > 5 And wordDeviation / longestWord.Length <= 0.2 Then
                        'LevenshteinDistance best
                        Return longestWord.Length - wordDeviation

                    Else
                        'Use ConsecutiveLetters
                        Dim successiveMatchCount = ConsecutiveLetters(shortestWord, longestWord).Key
                        If successiveMatchCount > 2 And successiveMatchCount / longestWord.Length >= 0.7 Then
                            Return longestWord.Length - successiveMatchCount
                        Else
                            Return 0
                        End If
                    End If
                End If
            End If
        Else
            Return 0
        End If

    End Function
    Public Function ShortestLongest(string1 As String, string2 As String) As String()

        string1 = If(string1, String.Empty)
        string2 = If(string2, String.Empty)
        Return If(string1.Length <= string2.Length, {string1, string2}, {string2, string1})

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
    Public Function PointsToRectangle(PointA As Point, PointB As Point) As Rectangle

        Dim upperLeftX As Integer = {PointA.X, PointB.X}.Min
        Dim upperLeftY As Integer = {PointA.Y, PointB.Y}.Min
        Dim bottomRightX As Integer = {PointA.X, PointB.X}.Max
        Dim bottomRightY As Integer = {PointA.Y, PointB.Y}.Max
        Return New Rectangle(New Point(upperLeftX, upperLeftY), New Size(bottomRightX - upperLeftX, bottomRightY - upperLeftY))

    End Function

#Region " ENUMS "
    Public Function RegexAbbreviatedMonthName(Optional grouped As Boolean = False) As String

        Dim months As New List(Of String)(EnumNames(GetType(AbbreviatedMonthName)).Except({"None"}))
        Dim monthString As String = Join(months.ToArray, "|")
        Return If(grouped, "(" & monthString & ")", monthString)

    End Function
    Public Function StringToAbbreviatedMonth(monthString As String) As AbbreviatedMonthName

        If monthString Is Nothing Then
            Return AbbreviatedMonthName.None
        Else
            Dim monthMatch As Match = Regex.Match(monthString, RegexAbbreviatedMonthName(True), RegexOptions.IgnoreCase)
            If monthMatch.Success Then
                Return ParseEnum(Of AbbreviatedMonthName)(monthMatch.Value)
            Else
                Return AbbreviatedMonthName.None
            End If
        End If

    End Function
    Public Enum AbbreviatedMonthName
        None = 0
        Jan = 1
        Feb = 2
        Mar = 3
        Apr = 4
        May = 5
        Jun = 6
        Jul = 7
        Aug = 8
        Sep = 9
        Oct = 10
        Nov = 11
        Dec = 12
    End Enum
    Public Function EnumNames(EnumItem As Object) As List(Of String)

        If EnumItem Is Nothing Then
            Return Nothing
        Else
            Return EnumNames(EnumItem.GetType)
        End If

    End Function
    Public Function EnumNames(EnumType As Type) As List(Of String)
        Try
            Return [Enum].GetNames(EnumType).ToList
        Catch ex As ArgumentException
            Return Nothing
        End Try
    End Function
    Public Function ParseEnum(Of T)(value As String) As T

        Dim enumValue As T
        For Each enumItem In EnumNames(GetType(T))
            If enumItem.ToUpperInvariant = value?.ToUpperInvariant Then
                enumValue = CType([Enum].Parse(GetType(T), enumItem, True), T)
            End If
        Next
        Return enumValue

    End Function
#End Region

    Public Function DataTypeToAlignment(valueType As Type) As HorizontalAlignment

        Select Case valueType
            Case GetType(Boolean), GetType(Byte), GetType(Short), GetType(Integer), GetType(Long), GetType(Date), GetType(DateAndTime), GetType(Image), GetType(Bitmap), GetType(Icon)
                Return HorizontalAlignment.Center

            Case GetType(Decimal), GetType(Double)
                Return HorizontalAlignment.Right

            Case Else
                Return HorizontalAlignment.Left

        End Select

    End Function
    Public Function DataTypeToFormat(valueType As Type) As String

        Select Case valueType

            Case GetType(Date)
                Dim CultureInfo = Threading.Thread.CurrentThread.CurrentCulture
                Return CultureInfo.DateTimeFormat.ShortDatePattern

            Case GetType(DateAndTime)
                Dim CultureInfo = Threading.Thread.CurrentThread.CurrentCulture
                Return CultureInfo.DateTimeFormat.FullDateTimePattern

            Case GetType(Decimal), GetType(Double)
                Return "C2"

            Case Else 'GetType(String), GetType(Byte), GetType(Short), GetType(Integer), GetType(Long), GetType(Boolean)
                Return String.Empty

        End Select

    End Function
    Public Function ContentAlignToStringFormat(alignString As String) As StringFormat

        Dim alignElements As New List(Of String)(Regex.Split(alignString, "(?=[A-Z])", System.Text.RegularExpressions.RegexOptions.None).Skip(1))
        If alignElements.Count = 2 Then
            'BottomLeft ... LineAlignment + Alignment
            Dim verticalAlignment As StringAlignment = If(alignElements.First = "Top", StringAlignment.Near, If(alignElements.First = "Middle", StringAlignment.Center, StringAlignment.Far))
            Dim horizontalAlignment As StringAlignment = If(alignElements.Last = "Left", StringAlignment.Near, If(alignElements.Last = "Center", StringAlignment.Center, StringAlignment.Far))
            Return New StringFormat With {
                .Alignment = horizontalAlignment,
                .LineAlignment = verticalAlignment}
        Else
            Return Nothing
        End If

    End Function
    Public Function ContentAlignToStringFormat(alignment As ContentAlignment) As StringFormat
        Return ContentAlignToStringFormat(alignment.ToString)
    End Function
    Public Function StringFormatToContentAlignString(formatString As StringFormat) As String

        If formatString Is Nothing Then
            Return Nothing
        Else
            Dim verticalString As String = If(formatString.LineAlignment = StringAlignment.Near, "Top", If(formatString.LineAlignment = StringAlignment.Center, "Middle", "Bottom"))
            Dim horizontalString As String = If(formatString.Alignment = StringAlignment.Near, "Left", If(formatString.Alignment = StringAlignment.Center, "Center", "Right"))
            Return verticalString & horizontalString
        End If

    End Function
    Public Function LiteralString(value As String) As String
        Dim sb As New StringBuilder(If(value, String.Empty))
        Return sb.ToString
    End Function

    Public Function DrawRoundedRectangle(Rect As Rectangle, Optional Corner As Integer = 10) As Drawing2D.GraphicsPath

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
    Public Function GetRoundedLine(points As PointF(), cornerRadius As Single) As Drawing2D.GraphicsPath

        If points Is Nothing Then
            Return Nothing
        Else
            Dim path As Drawing2D.GraphicsPath = New Drawing2D.GraphicsPath
            Dim previousEndPoint As PointF = PointF.Empty
            For i As Integer = 1 To points.Length - 1
                Dim startPoint As PointF = points(i - 1)
                Dim endPoint As PointF = points(i)

                If i > 1 Then
                    Dim cornerPoint As PointF = startPoint
                    LengthenLine(endPoint, startPoint, -cornerRadius)
                    Dim controlPoint1 As PointF = cornerPoint
                    Dim controlPoint2 As PointF = cornerPoint
                    LengthenLine(previousEndPoint, controlPoint1, -cornerRadius / 2)
                    LengthenLine(startPoint, controlPoint2, -cornerRadius / 2)
                    path.AddBezier(previousEndPoint, controlPoint1, controlPoint2, startPoint)
                End If

                If i + 1 < points.Length Then LengthenLine(startPoint, endPoint, -cornerRadius)
                path.AddLine(startPoint, endPoint)
                previousEndPoint = endPoint
            Next
            Return path
        End If

    End Function
    Public Sub LengthenLine(startPoint As PointF, ByRef endPoint As PointF, pixelCount As Single)

        If startPoint.Equals(endPoint) Then Return

        Dim dx As Double = endPoint.X - startPoint.X
        Dim dy As Double = endPoint.Y - startPoint.Y

        If dx = 0 Then
            If endPoint.Y < startPoint.Y Then
                endPoint.Y -= pixelCount
            Else
                endPoint.Y += pixelCount
            End If
        ElseIf dy = 0 Then
            If endPoint.X < startPoint.X Then
                endPoint.X -= pixelCount
            Else
                endPoint.X += pixelCount
            End If
        Else
            Dim length As Double = Math.Sqrt(dx * dx + dy * dy)
            Dim scale As Double = (length + pixelCount) / length
            dx *= scale
            dy *= scale
            endPoint.X = startPoint.X + Convert.ToSingle(dx)
            endPoint.Y = startPoint.Y + Convert.ToSingle(dy)
        End If

    End Sub
    Public Function DrawSpeechBubble(Rect As Rectangle) As Drawing2D.GraphicsPath

        Dim corner As Single = 22

        Dim Graphix As New System.Drawing.Drawing2D.GraphicsPath
        Dim ArcRect As New RectangleF(Rect.Location, New SizeF(corner, corner))

        Graphix.AddArc(ArcRect, 180, 90)
        Graphix.AddLine(Rect.X + CInt(corner / 2), Rect.Y,
                        Rect.X + Rect.Width - CInt(corner / 2), Rect.Y)
        ArcRect.X = Rect.Right - corner

        Graphix.AddArc(ArcRect, 270, 90)
        Graphix.AddLine(Rect.X + Rect.Width, Rect.Y + CInt(corner / 2),
                        Rect.X + Rect.Width, Rect.Y + Rect.Height - CInt(corner / 2))
        ArcRect.Y = Rect.Bottom - corner

        Graphix.AddArc(ArcRect, 0, 90)
        Graphix.AddLine(Rect.X + CInt(corner / 2), Rect.Y + Rect.Height,
                        Rect.X + Rect.Width - CInt(corner / 2), Rect.Y + Rect.Height)
        ArcRect.X = Rect.Left

        Graphix.AddArc(ArcRect, 90, 90)
        Graphix.AddLine(Rect.X - 18, Rect.Y + Rect.Height - (CInt(corner / 2)),
                        Rect.X, Rect.Y + CInt(corner / 2))

        Return Graphix

    End Function
    Friend Function SetOpacity(image As Image, opacity As Single) As Image

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
    Public Function RunIcon(runIndex As Long) As Icon

        Dim mod8 As Byte = CByte(runIndex Mod 8)
        Dim runningMan As Icon = Nothing
        If mod8 = 0 Then runningMan = My.Resources.r1
        If mod8 = 1 Then runningMan = My.Resources.r2
        If mod8 = 2 Then runningMan = My.Resources.r3
        If mod8 = 3 Then runningMan = My.Resources.r4
        If mod8 = 4 Then runningMan = My.Resources.r5
        If mod8 = 5 Then runningMan = My.Resources.r6
        If mod8 = 6 Then runningMan = My.Resources.r7
        If mod8 = 7 Then runningMan = My.Resources.r8
        Return runningMan

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
    Public Function Unzip(pathZip As String, folderExtract As String) As TriState

        If If(pathZip, String.Empty).Any And If(folderExtract, String.Empty).Any Then
            folderExtract = Path.GetFullPath(folderExtract)
            '// Ensures that the last character on the extraction path is the directory separator char.
            '// Without this, a malicious zip file could try to traverse outside of the expected extraction path.
            If Not folderExtract.EndsWith(Path.DirectorySeparatorChar.ToString(InvariantCulture), StringComparison.Ordinal) Then folderExtract += Path.DirectorySeparatorChar

            If File.Exists(pathZip) Then
                'The source zip file to read MUST exist, but the destination folder can be created
                If Not Directory.Exists(folderExtract) Then PathEnsureExists(folderExtract)
                Try
                    Using archive As ZipArchive = ZipFile.OpenRead(pathZip)
                        For Each entry As ZipArchiveEntry In archive.Entries
                            Dim destinationPath As String = Path.GetFullPath(Path.Combine(folderExtract, entry.FullName))
                            If destinationPath.StartsWith(folderExtract, StringComparison.Ordinal) Then
                                PathEnsureExists(destinationPath)
                                entry.ExtractToFile(destinationPath, True) '<=== Overwrites any existing files with same name
                            End If
                        Next
                    End Using
                    Return TriState.True
                Catch ex As InvalidDataException
                    Return TriState.False
                End Try
            Else
                Return TriState.UseDefault
            End If
        Else
            Return TriState.UseDefault
        End If

    End Function
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
    Public Function FileSize(info As FileInfo) As KeyValuePair(Of Double, String)

        If info Is Nothing Then
            Return Nothing
        Else
            Return (FileSize(info.Length))
        End If

    End Function
    Public Function FileSize(fileLength As Double) As KeyValuePair(Of Double, String)

        Dim sizes As String() = {"B", "KB", "MB", "GB", "TB"}
        Dim order As Integer = 0
        Do While fileLength >= 1024 And order < sizes.Length - 1
            order += 1
            fileLength /= 1024
        Loop
        Return New KeyValuePair(Of Double, String)(Math.Round(fileLength, 1), sizes(order))

    End Function
    Public Function ReadFilesInfo(FolderName As String) As Dictionary(Of FileInfo, String)

        If FolderName Is Nothing Then
            Return Nothing
        Else
            If Directory.Exists(FolderName) Then
                Return GetFiles(FolderName, ".txt").ToDictionary(Function(k) New FileInfo(k), Function(v) ReadText(v))
            Else
                Return Nothing
            End If
        End If

    End Function
    Public Function ReadFiles(FolderName As String) As Dictionary(Of String, String)

        If FolderName Is Nothing Then
            Return Nothing
        Else
            If Directory.Exists(FolderName) Then
                Return GetFiles(FolderName, ".txt").ToDictionary(Function(k) k, Function(v) ReadText(v))
            Else
                Return Nothing
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
                If kvp.Value = ExtensionNames.None Then
                    FilePathOrName &= ".txt"
                Else
                    'Is a filepath and has extension ( valid or not ) however does not exist at location
                    Return Nothing
                End If

            Else
                'Not a file format so could be Name only as ABC Or Name + Extension as ABC.txt...Assume to Desktop
                FilePathOrName = Desktop & "\" & FilePathOrName
                If kvp.Value = ExtensionNames.None Then FilePathOrName &= ".txt"

            End If
            'Try again after cleanup
            CanRead = IsFile(FilePathOrName) And File.Exists(FilePathOrName)
        End If
        If CanRead Then
            Dim Content As String = Nothing
            Try
                Using SR As New StreamReader(FilePathOrName)
                    Content = SR.ReadToEnd
                End Using
            Catch ex As IOException
                Return ex.Message
            End Try
            Return Content
        Else
            Return Nothing
        End If

    End Function
    Public Function IsFile(Source As String) As Boolean
        If Source Is Nothing Then
            Return False
        Else
            Return Regex.Match(Source, FilePattern, RegexOptions.IgnoreCase).Success
        End If
    End Function
    Public Function IsURL(address As String) As Boolean
        Return Regex.Match(address, "((([A-Za-z]{3,9}:(?:\/\/)?)(?:[\-;:&=\+\$,\w]+@)?[A-Za-z0-9\.\-]+|(?:www\.|[\-;:&=\+\$,\w]+@)[A-Za-z0-9\.\-]+)((?:\/[\+~%\/\.\w\-_]*)?\??(?:[\-\+=&;%@\.\w_]*)#?(?:[\.\!\/\\\w]*))?)", RegexOptions.IgnoreCase).Success
    End Function
    <Flags()> Public Enum ExtensionNames
        None
        Invalid
        Excel
        Text
        CommaSeparated
        PortableDocumentFormat
        SQL
        Unknown
    End Enum
    Public Function Settings() As List(Of System.Configuration.SettingsPropertyValue)
        Return (From spv In My.Settings.PropertyValues Select DirectCast(spv, System.Configuration.SettingsPropertyValue)).ToList
    End Function
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
    Public Function PathEnsureExists(filePath As String) As Boolean

        If filePath Is Nothing Then
            Return False
        Else
            If File.Exists(filePath) Then
                Return True
            Else
                Dim levels As New List(Of String)
                'C:\Users\SeanGlover\Desktop\PSRR\txts\FL\082020\
                Dim levelPath As String = String.Empty
                Dim fileLevels As New List(Of String)(Split(filePath, "\"))
                fileLevels.Remove(fileLevels.Last)
                fileLevels.ForEach(Sub(level)
                                       levelPath &= level & "\"
                                       Try
                                           If Not Directory.Exists(levelPath) Then Directory.CreateDirectory(levelPath)
                                           Do While Not Directory.Exists(levelPath)
                                           Loop
                                       Catch ex As UnauthorizedAccessException 'Fires on special folders - they can't be created
                                           levels.Add(level & "_" & ex.Message)
                                       Catch ex1 As IOException
                                           levels.Add(level & "_" & ex1.Message)
                                       End Try
                                   End Sub)
                If levels.Any Then Stop
                Return Not levels.Any
            End If
        End If

    End Function
    Public Function GetFiles(Path As String, Optional Extension As String = ".txt") As List(Of String)

        Return SafeWalk.EnumerateFiles(Path, "*" & Extension, SearchOption.AllDirectories).ToList

    End Function
    Public Function GetFiles(Folder As FileInfo, Optional Extension As String = ".txt") As List(Of String)

        If Folder Is Nothing Then
            Return Nothing
        Else
            Return GetFiles(Folder.FullName, Extension)
        End If

    End Function
    Public Function GetFiles(Path As String, Extension As String, Level As SearchOption) As List(Of String)
        Return SafeWalk.EnumerateFiles(Path, "*" & Extension, Level).ToList
    End Function
    Public Function GetFileNameExtension(Path As String) As KeyValuePair(Of String, ExtensionNames)

        Dim NameAndFilter As String = Split(Path, "\").Last
        Dim FileNameExtension As String() = Split(NameAndFilter, ".")
        Dim FileName As String = FileNameExtension.First

        If FileNameExtension.Count = 1 Then
            'Missing extension
            Return New KeyValuePair(Of String, ExtensionNames)(FileName, ExtensionNames.None)

        ElseIf FileNameExtension.Count = 2 Then
            Dim FileFilter As String = FileNameExtension.Last.ToUpperInvariant
            Select Case True
                Case FileFilter.StartsWith("XL", StringComparison.InvariantCulture)
                    Return New KeyValuePair(Of String, ExtensionNames)(FileName, ExtensionNames.Excel)

                Case FileFilter = "TXT"
                    Return New KeyValuePair(Of String, ExtensionNames)(FileName, ExtensionNames.Text)

                Case FileFilter = "CSV"
                    Return New KeyValuePair(Of String, ExtensionNames)(FileName, ExtensionNames.CommaSeparated)

                Case FileFilter = "SQL"
                    Return New KeyValuePair(Of String, ExtensionNames)(FileName, ExtensionNames.SQL)

                Case FileFilter = "PDF"
                    Return New KeyValuePair(Of String, ExtensionNames)(FileName, ExtensionNames.PortableDocumentFormat)

                Case Else
                    Return New KeyValuePair(Of String, ExtensionNames)(FileName, ExtensionNames.Unknown)

            End Select

        Else
            'Can never have 2 or more . in a filepath
            Return New KeyValuePair(Of String, ExtensionNames)(FileName, ExtensionNames.Invalid)
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
    Public Function ExcelToDataSet(excelPath As String) As DataSet

        If excelPath Is Nothing Then
            Return Nothing
        Else
            If File.Exists(excelPath) Then
                Try
                    Dim excelSet As DataSet = Nothing
                    Using stream As FileStream = File.Open(excelPath, FileMode.Open, FileAccess.Read)
                        Using excelReader As IExcelDataReader = ExcelReaderFactory.CreateOpenXmlReader(stream)
                            excelSet = excelReader.AsDataSet
                        End Using
                    End Using
                    Dim newSet As DataSet = excelSet.Clone
                    'The reader creates a table for each tab BUT with Column0, Column1, Column2 ... and the 1st row is used in the rows
                    For Each excelTable As DataTable In excelSet.Tables
                        If excelTable.AsEnumerable.Any Then
                            Dim firstRow As DataRow = excelTable.Rows(0)
                            For Each column As DataColumn In excelTable.Columns
                                column.ColumnName = firstRow(column).ToString
                            Next
                            excelTable.Rows.Remove(firstRow)
                            Dim newTable As DataTable = excelTable.Clone
                            For Each column As DataColumn In excelTable.Columns
                                Dim columnType As Type = GetDataType(column)
                                newTable.Columns(column.ColumnName).DataType = If(columnType Is GetType(DateAndTime), GetType(Date), columnType)
                            Next
                            For Each row In excelTable.AsEnumerable
                                newTable.Rows.Add(row.ItemArray)
                            Next
                        End If
                    Next
                    Return excelSet

                Catch ex As IOException
                    'If User has the file open
                    MsgBox(ex.Message)
                    Return Nothing

                End Try
            Else
                Return Nothing
            End If
        End If

    End Function
    Public Function DataColumnToList(Column As DataColumn) As List(Of Object)

        Dim Objects As New List(Of Object)
        If Column IsNot Nothing Then Objects = (From r In Column.Table.AsEnumerable Select r(Column)).ToList
        Return Objects

    End Function
    Public Function DataColumnToStrings(Column As DataColumn, Optional allowNulls As Boolean = True) As List(Of String)

        Dim Strings As New List(Of String)
        If Column IsNot Nothing Then
            If allowNulls Then
                Strings.AddRange(From r In Column.Table.AsEnumerable Select r(Column).ToString & String.Empty)
            Else
                Strings.AddRange(From r In Column.Table.AsEnumerable Select If(IsDBNull(r(Column)) Or IsNothing(r(Column)), String.Empty, r(Column).ToString & String.Empty))
            End If
        End If
        Return Strings

    End Function
#End Region

    Public Function EntitiesToString(html As String) As String

        html = Regex.Replace(html, "&rsquo;", "'", RegexOptions.None)
        html = Regex.Replace(html, "&rdquo;|&amp;quot;", """", RegexOptions.None)
        html = Regex.Replace(html, "&reg;", "®", RegexOptions.None)
        html = Regex.Replace(html, "&amp;trade;", "™", RegexOptions.None)
        html = Regex.Replace(html, "&amp; ", "& ", RegexOptions.None)
        Dim chrMatches = RegexMatches(html, "&#[0-9]{1,3};", RegexOptions.None)
        For Each chrMatch In chrMatches
            Dim shortChr As String = Chr(CInt(Regex.Match(chrMatch.Value, "[0-9]{1,3}", RegexOptions.None).Value))
            html = Regex.Replace(html, "&#[0-9]{1,3};", shortChr, RegexOptions.None)
        Next
        html = Replace(html, "&agrave;", "à")
        html = Replace(html, "&ndash;", "-")

        Return html

    End Function
    Public Function BalancingCharacters(inString As String, Optional leftSide As String = "(", Optional rightSide As String = Nothing, Optional isRegex As Boolean = False) As List(Of StringStartEnd)

        Dim strings As New List(Of StringStartEnd)

        Dim regexReserved As New List(Of String) From {"/", "\", "(", ")"} 'Start small, add at a later time
        Dim leftRight As New Dictionary(Of String, String) From {
            {"(", ")"},
            {"[", "]"},
            {"<", ">"},
            {"{", "}"},
            {"<table", "</table>"}
        }
        leftSide = If(leftSide, "(")
        rightSide = If(rightSide, leftRight(leftSide))
        Dim leftPattern As String = String.Empty
        Dim rightPattern As String = String.Empty

        If isRegex Then
            leftPattern = leftSide
            rightPattern = rightSide
        Else
            For Each letter In leftSide
                leftPattern &= If(regexReserved.Contains(letter), "\", String.Empty) & letter
            Next
            For Each letter In rightSide
                rightPattern &= If(regexReserved.Contains(letter), "\", String.Empty) & letter
            Next
        End If

        Dim leftMatches As New List(Of Match)(RegexMatches(inString, leftPattern, RegexOptions.Multiline).OrderByDescending(Function(m) m.Index))
        Dim rightMatches As New List(Of Match)(RegexMatches(inString, rightPattern, RegexOptions.Multiline).OrderBy(Function(m) m.Index))

        For Each leftMatch In leftMatches
            Dim rights As New List(Of Match)(From rm In rightMatches Where rm.Index > leftMatch.Index)
            If rights.Any Then
                Dim firstRight As Match = rights.First
                Dim leftStart As Integer = leftMatch.Index
                Dim rightEnd As Integer = firstRight.Index + firstRight.Length
                Dim stringLength As Integer = rightEnd - leftStart
                Dim leftrightString As String = inString.Substring(leftStart, stringLength)
                strings.Add(New StringStartEnd(leftrightString, leftStart, stringLength))
                rightMatches.Remove(firstRight)
            End If
        Next
        If strings.Any Then
            Return strings.OrderBy(Function(s) s.Start).ToList
        Else
            Return strings
        End If

    End Function

    Public Function NameToProperty(objectProperty As String) As System.Configuration.SettingsPropertyValue

        Dim mySettings As New List(Of System.Configuration.SettingsPropertyValue)(From pv In My.Settings.PropertyValues Select DirectCast(pv, System.Configuration.SettingsPropertyValue))
        Dim changedProperties As New List(Of System.Configuration.SettingsPropertyValue)(From ms In mySettings Where ms.Name.Contains(objectProperty))
        If changedProperties.Any Then
            Return changedProperties.First
        Else
            Return Nothing
        End If

    End Function

#Region " ENCRYPTION "
    Public Function Krypt(TextIn As String) As String
        Return Convert.ToBase64String(Encoding.Unicode.GetBytes(TextIn))
    End Function
    Public Function DeKrypt(TextIn As String) As String
        Try
            Return Encoding.Unicode.GetString(Convert.FromBase64String(TextIn))
        Catch ex As FormatException
            Return Nothing
        End Try
    End Function
#End Region
#Region " VALUE TYPES "
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬ D A T A   T Y P E   F R O M   S I N G U L A R   O B J E C T   I N S T A N C E
    ''' <summary>      
    ''' Logic Order - most common to least common ( faster ) <br/>      
    ''' 1] Nothing, 2] String, 3] Integers {a) Byte, b) Short, c) Integer, d) Long}, 4] Decimals 5] Boolean, 6] a) Date (midnight), b) DateAndTime, 7] a) Bitmap, b) Icon<br/>         
    ''' </summary> 
    Public Function GetDataType(objectValue As Object) As Type

        If IsDBNull(objectValue) Then
            Return Nothing
        Else
            If objectValue Is Nothing Then
                Return Nothing
            Else
                If objectValue.GetType Is GetType(String) Then
                    'IsString ( Most common )
                    If {"TRUE", "FALSE"}.Contains(objectValue.ToString.ToUpperInvariant) Then
                        Return GetType(Boolean)
                    Else
                        Return GetType(String)
                    End If
                Else
                    If IsNumeric(objectValue) Then
                        Dim longNumber As Long
                        Dim numberString As String = objectValue.ToString
                        If Long.TryParse(numberString, longNumber) Then
                            'Is a *** W H O L E *** Number between Byte and Long ... must work up from smallest object
                            Dim byteNumber As Byte
                            If Byte.TryParse(numberString, byteNumber) Then
                                'IsByte
                                Return GetType(Byte)
                            Else
                                Dim shortNumber As Short
                                If Short.TryParse(numberString, shortNumber) Then
                                    'IsShort
                                    Return GetType(Short)
                                Else
                                    Dim integerNumber As Integer
                                    If Integer.TryParse(numberString, integerNumber) Then
                                        'IsInteger
                                        Return GetType(Integer)
                                    Else
                                        'IsLong
                                        Return GetType(Long)
                                    End If
                                End If
                            End If
                        Else
                            'Is a *** D E C I M A L *** Number
                            Dim decimalNumber As Double
                            If Double.TryParse(numberString, decimalNumber) Then
                                'IsDecimal
                                Return GetType(Double)
                            Else
                                Dim booleanValue As Boolean
                                If Boolean.TryParse(numberString, booleanValue) Then
                                    'IsBoolean
                                    Return GetType(Boolean)
                                Else
                                    Return GetType(String)
                                End If
                            End If
                        End If
                    Else
                        'Could be Dates
                        If objectValue.GetType Is GetType(Date) Then
                            Dim dateValue As Date = DirectCast(objectValue, Date)
                            If dateValue.Date = dateValue Then
                                'IsDate ( No Hours, Minutes, Seconds )
                                Return GetType(Date)
                            Else
                                'IsDateTime ( Hours, Minutes, Seconds )
                                Return GetType(DateAndTime)
                            End If
                        Else
                            Dim imageValue As Bitmap = TryCast(objectValue, Bitmap)
                            If imageValue Is Nothing Then
                                Dim iconValue As Icon = TryCast(objectValue, Icon)
                                If iconValue Is Nothing Then
                                    'IsString ( Default )
                                    Return GetType(String)
                                Else
                                    'IsIcon
                                    Return GetType(Icon)
                                End If
                            Else
                                'IsImage
                                Return GetType(Image)
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End Function
    Public Function GetDataType(Value As String, Optional Test As Boolean = False) As Type

        If Value Is Nothing Then
            Return GetType(String)

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
                Dim _Boolean As Boolean
                If Boolean.TryParse(Value, _Boolean) Or Value.ToUpperInvariant = "TRUE" Or Value.ToUpperInvariant = "FALSE" Then
                    Return _Boolean.GetType

                Else
                    Dim _Date As Date
                    Dim dateFormats() As String = {
                    "M/d/yyyy",
                    "M/d/yyyy h:mm",
                    "M/d/yyyy h:mm:ss",
                    "M/d/yyyy h:mm:ss tt",
                    "yyyy-M-d h:mm:ss tt"
                    } '2019-11-06 12:00:00 AM
                    'Date.Parse("2020-01-21T18:25:24Z", New Globalization.CultureInfo("en-US"), Globalization.DateTimeStyles.AdjustToUniversal)
                    If Date.TryParse(Value, CultureInfo.CurrentCulture, DateTimeStyles.AdjustToUniversal, _Date) Or Date.TryParseExact(Value, dateFormats, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, _Date) Then
                        If _Date.Date = _Date Then
                            If Test Then Stop
                            Return _Date.GetType
                        Else
                            If Test Then Stop
                            Return GetType(DateAndTime)
                        End If
                    Else
                        'Some objects can not be converted in the ToString Function ... they only show as the object name
                        If Value.Contains("Drawing.Bitmap") Or Value.Contains("Drawing.Image") Then
                            Return GetType(Image)

                        ElseIf Value.Contains("Drawing.Icon") Then
                            Return GetType(Icon)

                        Else
                            Return GetType(String)

                        End If
                    End If
                End If
            End If
        End If

    End Function
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬ D A T A   T Y P E   F R O M   O B J E C T   C O L L E C T I O N
    Public Function GetDataType(Column As DataColumn) As Type

        Dim values As New List(Of Object)(DataColumnToList(Column))
        Dim valuesType As Type = GetDataType(values) 'TESTING:GetDataType with a Boolean Parameter ... If(Column Is Nothing, String.Empty, Column.ColumnName).EndsWith("_DATE", StringComparison.InvariantCulture)
        Return valuesType

    End Function
    Public Function GetDataType(Values As List(Of Object), Optional Test As Boolean = False) As Type

        If Values Is Nothing Then
            Return Nothing
        Else
            Dim nonNull As New List(Of Object)(From v In Values Where Not (IsDBNull(v) Or IsNothing(v))) 'Null values say nothing about potential Type
            If nonNull.Any Then

                If Test Then Stop

                Dim Types As New List(Of Type)(From nn In nonNull Select GetDataType(nn))
                Dim aggregateType As Type = GetDataType(Types.Distinct)
                If aggregateType Is GetType(String) Then
                    'Switch String to Boolean ???, Yes if all non-null values are "True" and "False"
                    Dim tfDictionary As New Dictionary(Of String, List(Of String)) From {
                {"TRUE", New List(Of String)},
                {"FALSE", New List(Of String)}
                }
                    Dim ynDictionary As New Dictionary(Of String, List(Of String)) From {
                {"Y", New List(Of String)},
                {"N", New List(Of String)}
                }
                    For Each value In nonNull
                        Dim upperString As String = value.ToString.ToUpperInvariant
                        If tfDictionary.ContainsKey(upperString) Then tfDictionary(upperString).Add(upperString)
                        If ynDictionary.ContainsKey(upperString) Then ynDictionary(upperString).Add(upperString)
                    Next
                    Dim tfAll As Boolean = tfDictionary("TRUE").Any And tfDictionary("FALSE").Any And tfDictionary("TRUE").Count + tfDictionary("FALSE").Count = nonNull.Count
                    Dim ynAll As Boolean = ynDictionary("Y").Any And ynDictionary("N").Any And ynDictionary("Y").Count + ynDictionary("N").Count = nonNull.Count
                    Return If(tfAll Or ynAll, GetType(Boolean), GetType(String))
                Else
                    Dim kvp = Column.Get_kvpFormat(aggregateType)
                    If kvp.Key = Column.TypeGroup.Integers Then
                        'Switch whole numbers to Boolean ???, Yes if all non-null values are 0 and 1 ... Maybe add a minimum count restriction of 3 each??? 0.Count=3 and 1.Count=3
                        Dim booleanDictionary As New Dictionary(Of Long, List(Of Long)) From {
                {1, New List(Of Long)},
                {0, New List(Of Long)}
                }
                        For Each value In nonNull
                            Dim longTrueFalse As Long = CLng(value)
                            If booleanDictionary.ContainsKey(longTrueFalse) Then booleanDictionary(longTrueFalse).Add(longTrueFalse)
                        Next
                        Return If(booleanDictionary(0).Any And booleanDictionary(1).Any And booleanDictionary(0).Count + booleanDictionary(1).Count = nonNull.Count, GetType(Boolean), aggregateType)
                    Else
                        Return aggregateType
                    End If
                End If
            Else
                Return Nothing
            End If
        End If

    End Function
    Public Function GetDataType(Types As List(Of String), Optional test As Boolean = False) As Type

        If Types Is Nothing Then
            Return Nothing
        Else
            If test Then
                'If test Then Stop
                'If Types.Contains("D081") Then Stop
            End If
            Dim typeList = From t In Types Select GetDataType(t)
            Dim aggregateType As Type = GetDataType(typeList.ToList)
            Return aggregateType
        End If

    End Function
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬ D A T A   T Y P E   F R O M   T Y P E   C O L L E C T I O N
    Public Function GetDataType(Types As List(Of Type), Optional testing As Boolean = False) As Type

        If Types Is Nothing Then
            Return Nothing
        Else
            Dim distinctTypes As New List(Of Type)((From t In Types Where Not (t Is Nothing Or IsDBNull(t))).Distinct)
            If distinctTypes.Any Then
                Dim typeCount As Integer = distinctTypes.Count
                If testing Then Stop

                If typeCount = 1 Then
                    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■  ONLY 1 TYPE, RETURN IT
                    Return distinctTypes.First

                Else
                    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■  MULTIPLE TYPES - CHOOSE BEST FIT ex) Date + DateAndTime = DateAndTime, Byte + Short = Short
                    If distinctTypes.Intersect({GetType(Date), GetType(DateAndTime)}).Count = typeCount Then
                        Return GetType(DateAndTime)
                    Else
                        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■  NUMERIC
                        If distinctTypes.Intersect({GetType(Byte), GetType(Short), GetType(Integer), GetType(Long), GetType(Double), GetType(Decimal)}).Count = typeCount Then
                            If distinctTypes.Intersect({GetType(Byte), GetType(Short), GetType(Integer), GetType(Long)}).Count = typeCount Then
                                '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■  MIX OF INTEGER ... DESCEND IN SIZE TO GET LARGEST NECESSARY
                                If distinctTypes.Contains(GetType(Long)) Then
                                    Return GetType(Long)
                                Else
                                    If distinctTypes.Contains(GetType(Integer)) Then
                                        Return GetType(Integer)
                                    Else
                                        If distinctTypes.Contains(GetType(Short)) Then
                                            Return GetType(Short)
                                        Else
                                            Return GetType(Byte)
                                        End If
                                    End If
                                End If
                            Else
                                '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■  COULD BE MIX OF INTEGER, DECIMAL, DOUBLE
                                Return GetType(Double)
                            End If
                        Else
                            If distinctTypes.Intersect({GetType(Image), GetType(Bitmap)}).Count = typeCount Then
                                '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■  IMAGE / BITMAP
                                Return GetType(Bitmap)
                            Else
                                If distinctTypes.Intersect({GetType(Image), GetType(Bitmap), GetType(Icon)}).Any Then
                                    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■  IMAGES DON'T MIX WITH OTHER VALUES AS THEY CAN'T REPRESENTED IN A TEXT FORM
                                    Return GetType(Object)
                                Else
                                    Return GetType(String)
                                End If
                            End If
                        End If
                    End If
                    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■  STRING AS DEFAULT
                    Return GetType(String)
                End If
            Else
                Return Nothing
            End If
        End If

    End Function
    Public Function GetDataType(Types As IEnumerable(Of Type)) As Type

        If Types Is Nothing Then
            Return Nothing
        Else
            Return GetDataType(Types.ToList)
        End If

    End Function
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Function ValueToField(Value As Object) As String

        If Value Is Nothing Then
            Return Nothing
        Else
            Return ValueToField(Value, Value.GetType)
        End If

    End Function
    Public Function ValuesToFields(Values As Object()) As String

        If Values Is Nothing Then
            Return Nothing
        Else
            Dim Items As New List(Of String)
            For Each Value In Values
                Items.Add(ValueToField(Value, GetDataType(Value)))
            Next
            Return "(" & Join(Items.ToArray, ",") & ")"
        End If

    End Function
    Public Function ValueToField(Value As Object, ValueType As Type) As String

        If Value Is Nothing Then
            Return Nothing
        Else
            Select Case ValueType
                Case GetType(String)
                    Return Join({"'", Value.ToString, "'"}, String.Empty)

                Case GetType(Date)
                    Dim DateValue As Date = DirectCast(Value, Date)
                    If DateValue.TimeOfDay = New TimeSpan(0) Then
                        Return DateToDB2Date(DateValue)
                    Else
                        Return DateToDB2Timestamp(DateValue)
                    End If

                Case Else
                    Return Value.ToString

            End Select
        End If

    End Function
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
#End Region
    Public Enum HandlerAction
        Add
        Remove
    End Enum
End Module

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

                Return New Size(2 * BorderThickness + ColumnWidths + 6, 2 * BorderThickness + RowHeights + 6)
            Else
                Return Nothing
            End If

        End Function
        Public Function GetColumns(TLP As TableLayoutPanel) As Dictionary(Of Integer, List(Of Control))

            If TLP Is Nothing Then
                Return Nothing

            Else
                Dim Columns As New Dictionary(Of Integer, List(Of Control))
                For Each Item As Control In TLP.Controls
                    Dim xy As TableLayoutPanelCellPosition = TLP.GetCellPosition(Item)
                    If Not Columns.Keys.Contains(xy.Column) Then Columns.Add(xy.Column, New List(Of Control))
                    Columns(xy.Column).Add(Item)
                Next
                Return Columns
            End If

        End Function
        Public Function GetRows(TLP As TableLayoutPanel) As Dictionary(Of Integer, List(Of Control))

            If TLP Is Nothing Then
                Return Nothing

            Else
                Dim Rows As New Dictionary(Of Integer, List(Of Control))
                For Each Item As Control In TLP.Controls
                    Dim xy As TableLayoutPanelCellPosition = TLP.GetCellPosition(Item)
                    If Not Rows.Keys.Contains(xy.Row) Then Rows.Add(xy.Row, New List(Of Control))
                    Rows(xy.Row).Add(Item)
                Next
                Return Rows
            End If

        End Function
        Public Sub SetSize(TLP As TableLayoutPanel, Optional distributeColumnWidths As Boolean = False)

            If TLP IsNot Nothing Then
                TLP.Size = GetSize(TLP)
                If distributeColumnWidths Then DistributeWidths(TLP)
            End If

        End Sub
        Public Sub DistributeWidths(TLP As TableLayoutPanel)

            If TLP IsNot Nothing Then
                Dim columnsWidth As Single = 0
                Dim absoluteCount As Integer = 0
                For Each column As ColumnStyle In TLP.ColumnStyles
                    If column.SizeType = SizeType.Absolute Then
                        columnsWidth += column.Width
                        absoluteCount += 1
                    End If
                Next
                If absoluteCount = TLP.ColumnStyles.Count Then 'Can change ... percentage is a pain
                    If TLP.Width > columnsWidth Then 'Should be some blank space
                        Dim deltaWidth As Integer = TLP.Width - CInt(columnsWidth) 'Extra
                        TLP.ColumnStyles(TLP.ColumnStyles.Count - 1).Width += deltaWidth
                    End If
                End If
            End If

        End Sub
    End Module
End Namespace

Public NotInheritable Class SafeWalk
    Public Sub New()
    End Sub
    Public Shared Function EnumerateFiles(Path As String, SearchPattern As String, SearchOpt As SearchOption) As IEnumerable(Of String)

        Try
            Dim di As DirectoryInfo = New DirectoryInfo(Path)
            Dim files As FileInfo() = di.GetFiles(SearchPattern, SearchOpt)
            Return files.Select(Function(f) f.FullName)

        Catch ex As DirectoryNotFoundException
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
    Private Function CalculateAbsoluteCoordinateX(x As Integer) As Integer
        Return CType((x * 65536) / NativeMethods.GetSystemMetrics(SystemMetric.SMxCXSCREEN), Integer)
    End Function
    Private Function CalculateAbsoluteCoordinateY(y As Integer) As Integer
        Return CType((y * 65536) / NativeMethods.GetSystemMetrics(SystemMetric.SMxCYSCREEN), Integer)
    End Function
    Public Sub ClickLeftMouseButton(Location As Point)
        ClickLeftMouseButton(Location.X, Location.Y)
    End Sub
    Public Sub ClickLeftMouseButton(x As Integer, y As Integer)

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
    Public Sub ClickRightMouseButton(Location As Point)
        ClickRightMouseButton(Location.X, Location.Y)
    End Sub
    Public Sub ClickRightMouseButton(x As Integer, y As Integer)

        Dim MouseInput As INPUT = New INPUT With {
            .type = SendInputEventType.InputMouse
        }
        With MouseInput
            .mkhi.mi.dx = CalculateAbsoluteCoordinateX(x)
            .mkhi.mi.dy = CalculateAbsoluteCoordinateY(y)
            .mkhi.mi.mouseData = 0
            .mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTFxMOVE Or MouseEventFlags.MOUSEEVENTFxABSOLUTE
            Dim unused1 = NativeMethods.SendInput(1, MouseInput, Marshal.SizeOf(New INPUT()))
            .mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTFxRIGHTDOWN
            Dim unused2 = NativeMethods.SendInput(1, MouseInput, Marshal.SizeOf(New INPUT()))
            .mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTFxRIGHTUP
            Dim unused3 = NativeMethods.SendInput(1, MouseInput, Marshal.SizeOf(New INPUT()))
        End With

    End Sub
    Public Sub MoveMouse(Location As Point)
        MoveMouse(Location.X, Location.Y)
    End Sub
    Public Sub MoveMouse(x As Integer, y As Integer)

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
    Public Sub KeyPress(keyCode As Keys)

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
Public Module html
    Friend Event ElementWatched(sender As TimeSpan, e As List(Of HtmlElement))
    Private ReadOnly Property Document As HtmlDocument
    Private ReadOnly Property ElementIdName As String
    Private ReadOnly ElementStopWatch As New Stopwatch
    Private ReadOnly Property StopWatchLimit As Integer
    Private WithEvents ElementTimer As New Timer With {.Interval = 100}
    Public Enum InputType
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
        button
        checkbox
        color
        email
        file
        hidden
        image
        month
        password
        radio
        range
        reset
        search
        submit
        tel
        text
        time
        url
        week
    End Enum
    Public Enum SubmitType
        None
        Click
        Enter
    End Enum
    Public Function ElementInputSubmitType(Element As HtmlElement) As KeyValuePair(Of InputType, SubmitType)

        Dim it As InputType

        If Element Is Nothing Then
            Return New KeyValuePair(Of InputType, SubmitType)(it, SubmitType.None)

        Else
            Dim InputMatch As Match = Regex.Match(Element.OuterHtml, "(?<= type="")[^""]{1,}", RegexOptions.IgnoreCase)
            If InputMatch.Success Then
                it = ParseEnum(Of InputType)(InputMatch.Value)
            Else
                InputMatch = Regex.Match(Element.OuterHtml, "onclick=", RegexOptions.IgnoreCase)
                it = If(InputMatch.Success, InputType.button, InputType.None)
            End If
        End If

        Dim st As SubmitType = If(it = InputType.text Or it = InputType.password, SubmitType.Enter, SubmitType.Click)
        Return New KeyValuePair(Of InputType, SubmitType)(it, st)

    End Function
    Public Function ElementByRegex(Document As HtmlDocument, SearchValue As String, Optional searchBy As String = "id", Optional searchTag As String = Nothing) As HtmlElement

        Dim elements As List(Of HtmlElement) = ElementsByRegex(Document, SearchValue, searchBy, searchTag)
        Return If(elements.Any, elements.First, Nothing)

    End Function
    Public Function ElementsByRegex(Document As HtmlDocument, searchValue As String, Optional searchBy As String = "id", Optional searchTag As String = Nothing) As List(Of HtmlElement)

        searchValue = If(searchValue, String.Empty)
        Dim elements As New List(Of HtmlElement)

        If Document IsNot Nothing Then
            If Document.Body IsNot Nothing Then
                searchBy = If(searchBy, "id")
                If searchBy.ToUpperInvariant = "ID" Then
                    elements.Add(Document.GetElementById(searchValue))

                ElseIf searchBy.ToUpperInvariant = "TEXT" Then
                    For Each windowFrame As HtmlWindow In Document.Window.Frames
                        Try
                            Dim frameDocument As HtmlDocument = windowFrame.Document
                            For Each frameForm As HtmlElement In frameDocument.Forms
                                For Each element As HtmlElement In frameForm.All
                                    If If(element.InnerText, String.Empty).ToUpperInvariant = searchValue.ToUpperInvariant Then
                                        elements.Add(element)
                                    End If
                                Next
                            Next
                        Catch ex As UnauthorizedAccessException
                            If searchValue = "Search" Then Stop
                        End Try
                    Next
                    If searchValue = "Integrated Receivables Output List" And elements.Any Then Stop
                Else
                    Dim elementsAll As New Dictionary(Of String, HtmlElement)
                    For Each element In From e In Document.All Select DirectCast(e, HtmlElement)
                        Dim formatArray As Object() = {elementsAll.Count, element.TagName}
                        Dim elementKey As String = String.Format(InvariantCulture, "{0:000} {1:}", formatArray)
                        elementsAll.Add(elementKey, element)
                    Next
                    elements.AddRange(From ea In elementsAll.Values Where ea.TagName = searchBy Select ea)
                    If elements.Any Then
                        Dim values As New List(Of HtmlElement)
                        If searchBy.ToUpperInvariant = "NAME" Then
                            values.AddRange(From e In elements Where e.Name = searchValue)
                        Else
                            'href="View ATL-534151"
                            values.AddRange(From e In elements Where Regex.Match(If(e.OuterHtml, String.Empty), searchValue, RegexOptions.IgnoreCase).Success)
                        End If
                        elements = values
                    End If
                End If
            End If
        End If
        If searchTag Is Nothing Then
            Return elements
        ElseIf searchTag.Any Then
            Return elements.Where(Function(e) If(e.TagName, String.Empty).ToUpperInvariant = searchTag.ToUpperInvariant).ToList
        Else
            Return elements
        End If

    End Function
    Public Function ElementsByTag(Document As HtmlDocument) As Dictionary(Of String, List(Of HtmlElement))

        Dim Elements As New Dictionary(Of String, List(Of HtmlElement))
        If Document IsNot Nothing Then
            Dim OuterHtml As String = Document.Body.OuterHtml
            Dim Tags = New List(Of String) From {"a", "body", "br", "div", "Form", "h1", "h2", "h3", "h4", "head", "html", "iframe", "img", "input", "li", "link", "meta", "ol", "OptionOn", "p", "script", "select", "span", "style", "table", "th", "td", "textarea", "title", "tr", "ul"}
            For Each Tag In Tags
                For Each Element As HtmlElement In Document.GetElementsByTagName(Tag)
                    If Not Elements.ContainsKey(Tag) Then Elements.Add(Tag, New List(Of HtmlElement))
                    Elements(Tag).Add(Element)
                Next
            Next
        End If
        Return Elements

    End Function
    Public Function ElementsAll(Document As HtmlDocument) As List(Of HtmlElement)

        Dim ElementsDictionary = ElementsByTag(Document)
        Dim AllElements As New List(Of HtmlElement)
        For Each Tag In ElementsDictionary.Keys
            AllElements.AddRange(ElementsDictionary(Tag))
        Next
        Return AllElements

    End Function
    Public Function ElementsByKeyText(Document As HtmlDocument, SearchValue As String) As List(Of HtmlElement)

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
    Public Function SubmitForm(Document As HtmlDocument, Element As HtmlElement) As HtmlElement

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
    Public Sub ElementWatch(WebDocument As HtmlDocument, IdName As String, Optional Timeout As Integer = 10)
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
Public NotInheritable Class PropertyConverter
    Inherits TypeConverter
    Public Sub New()
    End Sub
    Public Overloads Overrides Function CanConvertFrom(context As ITypeDescriptorContext, sourceType As Type) As Boolean
        If sourceType?.Equals(GetType(String)) Then
            Return True
        Else
            Return MyBase.CanConvertFrom(context, sourceType)
        End If
    End Function
    Public Overloads Overrides Function CanConvertTo(context As ITypeDescriptorContext, destinationType As Type) As Boolean
        If destinationType?.Equals(GetType(String)) Then
            Return True
        Else
            Return MyBase.CanConvertTo(context, destinationType)
        End If
    End Function
    Public Overloads Overrides Function ConvertTo(context As ITypeDescriptorContext, culture As Globalization.CultureInfo, value As Object, destinationType As Type) As Object
        If destinationType?.Equals(GetType(String)) Then
            Return value?.ToString()
        Else
            Return MyBase.ConvertTo(context, culture, value, destinationType)
        End If
    End Function
    Public Overloads Overrides Function GetPropertiesSupported(context As ITypeDescriptorContext) As Boolean
        Return True
    End Function
    Public Overloads Overrides Function GetProperties(context As ITypeDescriptorContext, value As Object, Attribute() As Attribute) As PropertyDescriptorCollection
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
    Public Shared Function CreateCursor(bmp As Bitmap, xHotspot As Integer, yHotspot As Integer) As Cursor

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

    Public Shared Sub FixBrowserVersion(appName As String)
        FixBrowserVersion(appName, GetEmbVersion())
    End Sub
    ' End Sub FixBrowserVersion
    Public Shared Sub FixBrowserVersion(appName As String, ieVer As Integer)
        FixBrowserVersion_Internal("HKEY_LOCAL_MACHINE", appName & ".exe".ToString(InvariantCulture), ieVer)
        FixBrowserVersion_Internal("HKEY_CURRENT_USER", appName & ".exe".ToString(InvariantCulture), ieVer)
        FixBrowserVersion_Internal("HKEY_LOCAL_MACHINE", appName & ".vshost.exe".ToString(InvariantCulture), ieVer)
        FixBrowserVersion_Internal("HKEY_CURRENT_USER", appName & ".vshost.exe".ToString(InvariantCulture), ieVer)
    End Sub
    ' End Sub FixBrowserVersion
    Private Shared Sub FixBrowserVersion_Internal(root As String, appName As String, ieVer As Integer)
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
Public NotInheritable Class CustomRenderer
    Inherits ToolStripProfessionalRenderer
    'https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.toolstripprofessionalrenderer?view=netcore-3.1#events
    Public Enum ColorTheme
        Brown
        Black
        Green
        Blue
        Red
        Gray
        Yellow
        Purple
    End Enum
    Private ReadOnly Property ThemeColor As Color
    Public Property Theme As ColorTheme
    Private Const UnderlineAlpha As Byte = 128
    Public Sub New()

        ThemeColor = Color.FromArgb(103, 71, 205)
        Select Case Theme
            Case ColorTheme.Purple
                ThemeColor = Color.FromArgb(103, 71, 205)
            Case ColorTheme.Yellow
                ThemeColor = Color.Yellow
        End Select

    End Sub
    Protected Overrides Sub OnRenderToolStripBackground(e As ToolStripRenderEventArgs)

        'Entire Toolstrip
        If e IsNot Nothing Then
            'e.ToolStrip.Items.OfType(Of ToolStripItem).ToList().ForEach(Sub(Item As ToolStripItem)
            '                                                                If Item.Image Is Nothing Then
            '                                                                    Using backBrush As New SolidBrush(Color.FromArgb(UnderlineAlpha, Color.WhiteSmoke))
            '                                                                        e.Graphics.FillRectangle(backBrush, Item.ContentRectangle)
            '                                                                    End Using
            '                                                                End If
            '                                                                Dim underlineBounds As New Rectangle(Item.ContentRectangle.X, Item.ContentRectangle.Height - 6, Item.ContentRectangle.Width, 6)
            '                                                                Using backBrush As New SolidBrush(Color.FromArgb(UnderlineAlpha, ThemeColor))
            '                                                                    e.Graphics.FillRectangle(backBrush, underlineBounds)
            '                                                                End Using
            '                                                            End Sub)
        End If

    End Sub
    Protected Overrides Sub OnRenderImageMargin(e As ToolStripRenderEventArgs)

        MyBase.OnRenderImageMargin(e)
        If e IsNot Nothing Then
            Dim MarginWidth As Integer = e.AffectedBounds.Width
            e.ToolStrip.Items.OfType(Of ToolStripControlHost).ToList().ForEach(Sub(Item As ToolStripControlHost)
                                                                                   If Item.Image IsNot Nothing Then
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
                                                                               End Sub)
        End If

    End Sub
    Protected Overrides Sub OnRenderButtonBackground(e As ToolStripItemRenderEventArgs)

        If e IsNot Nothing Then
            If e.Item.Selected Then
                '/// Left Mouse Down
                If e.Item.Image Is Nothing Then
                    Using backBrush As New SolidBrush(Color.FromArgb(64, Color.WhiteSmoke))
                        e.Graphics.FillRectangle(backBrush, e.Item.ContentRectangle)
                    End Using
                End If
                Dim underlineBounds As New Rectangle(e.Item.ContentRectangle.X, e.Item.ContentRectangle.Height - 6, e.Item.ContentRectangle.Width, 6)
                Using backBrush As New SolidBrush(Color.FromArgb(UnderlineAlpha, ThemeColor))
                    e.Graphics.FillRectangle(backBrush, underlineBounds)
                End Using
            Else
                'If e.Item.Image Is Nothing Then
                '    Using backBrush As New SolidBrush(Color.FromArgb(64, Color.WhiteSmoke))
                '        e.Graphics.FillRectangle(backBrush, e.Item.ContentRectangle)
                '    End Using
                'End If
                'Dim underlineBounds As New Rectangle(e.Item.ContentRectangle.X, e.Item.ContentRectangle.Height - 6, e.Item.ContentRectangle.Width, 6)
                'Using backBrush As New SolidBrush(Color.FromArgb(UnderlineAlpha, ThemeColor))
                '    e.Graphics.FillRectangle(backBrush, underlineBounds)
                'End Using
            End If
        End If

    End Sub
    Protected Overrides Sub OnRenderDropDownButtonBackground(e As ToolStripItemRenderEventArgs)

        If e IsNot Nothing Then
            If e.Item.Selected Then
                '/// Left Mouse Down
                If e.Item.Image Is Nothing Then
                    Using backBrush As New SolidBrush(Color.FromArgb(UnderlineAlpha, Color.WhiteSmoke))
                        e.Graphics.FillRectangle(backBrush, e.Item.ContentRectangle)
                    End Using
                End If
                Dim underlineBounds As New Rectangle(e.Item.ContentRectangle.X, e.Item.ContentRectangle.Height - 6, e.Item.ContentRectangle.Width, 6)
                Using backBrush As New SolidBrush(Color.FromArgb(UnderlineAlpha, ThemeColor))
                    e.Graphics.FillRectangle(backBrush, underlineBounds)
                End Using
            Else

            End If
        End If

    End Sub
    Protected Overrides Sub OnRenderItemBackground(e As ToolStripItemRenderEventArgs)
        MyBase.OnRenderItemBackground(e)
    End Sub
    Protected Overrides Sub OnRenderMenuItemBackground(e As ToolStripItemRenderEventArgs)

        If e IsNot Nothing Then
            Using Brush As New SolidBrush(e.Item.BackColor)
                e.Graphics.FillRectangle(Brush, e.Item.ContentRectangle)
            End Using
            If e.Item.Selected Then
                Dim underlineBounds As New Rectangle(e.Item.ContentRectangle.X, e.Item.ContentRectangle.Height - 6, e.Item.ContentRectangle.Width, 6)
                Using backBrush As New SolidBrush(Color.FromArgb(UnderlineAlpha, ThemeColor))
                    e.Graphics.FillRectangle(backBrush, underlineBounds)
                End Using
            End If
        End If

    End Sub
End Class
Public NotInheritable Class ThreadHelperClass
    Delegate Sub SetToolPropertyCallback(tsi As ToolStripItem, n As String, v As Object)
    Delegate Sub SetPropertyCallback(c As Control, n As String, v As Object)
    Delegate Sub GetPropertyCallback(c As Control, n As String)
    Public Shared Sub SetSafeControlPropertyValue(Item As Control, PropertyName As String, PropertyValue As Object)

        If Item IsNot Nothing Then
            Try
                If Item.InvokeRequired Then
                    Dim d As SetPropertyCallback = New SetPropertyCallback(AddressOf SetSafeControlPropertyValue)
                    Item.Invoke(d, New Object() {Item, PropertyName, PropertyValue})

                Else
                    Dim t As Type = Item.GetType
                    Dim pi As PropertyInfo = t.GetProperty(PropertyName)
                    pi.SetValue(Item, PropertyValue)

                End If
            Catch ex As ObjectDisposedException
            Catch ex As InvalidAsynchronousStateException
            Catch ex As TargetInvocationException
            End Try
        End If

    End Sub
    Public Shared Sub SetSafeToolStripItemPropertyValue(Item As ToolStripItem, PropertyName As String, PropertyValue As Object)

        If Item IsNot Nothing Then
            Try
                If Item.Owner.InvokeRequired Then
                    Dim d As SetToolPropertyCallback = New SetToolPropertyCallback(AddressOf SetSafeToolStripItemPropertyValue)
                    Item.Owner.Invoke(d, New Object() {Item, PropertyName, PropertyValue})

                Else
                    Dim t As Type = Item.GetType
                    Dim pi As PropertyInfo = t.GetProperty(PropertyName)
                    pi.SetValue(Item, PropertyValue)
                End If
            Catch ex As ObjectDisposedException
            Catch ex As InvalidAsynchronousStateException
            Catch ex As TargetInvocationException
            End Try
        End If

    End Sub
    Public Shared Function GetSafeControlPropertyValue(Item As Control, PropertyName As String) As Object

        If Item Is Nothing Then
            Return Nothing
        Else
            Try
                Dim t As Type = Item.GetType
                If Item.InvokeRequired Then
                    Dim d As GetPropertyCallback = New GetPropertyCallback(AddressOf GetSafeControlPropertyValue)
                    Try
                        Return Item.Invoke(d, New Object() {Item, PropertyName})
                    Catch ex As TargetInvocationException
                    End Try
                    Return Nothing

                Else
                    Dim pi As PropertyInfo = t.GetProperty(PropertyName)
                    Try
                        Return pi.GetValue(Item)
                    Catch ex As TargetInvocationException
                        Return Nothing
                    End Try

                End If
            Catch ex As ObjectDisposedException
                Return Nothing
            End Try
        End If

    End Function
End Class
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
    Public Overrides Function GetHashCode() As Integer
        Return CbSize.GetHashCode Xor FMask.GetHashCode Xor NMin.GetHashCode Xor NPage.GetHashCode Xor NPos.GetHashCode Xor NTrackPos.GetHashCode
    End Function
    Public Overloads Function Equals(other As SCROLLINFO) As Boolean Implements IEquatable(Of SCROLLINFO).Equals
        Return CbSize = other.CbSize AndAlso FMask = other.FMask AndAlso NMin = other.NMin
    End Function
    Public Shared Operator =(value1 As SCROLLINFO, value2 As SCROLLINFO) As Boolean
        Return value1.Equals(value2)
    End Operator
    Public Shared Operator <>(value1 As SCROLLINFO, value2 As SCROLLINFO) As Boolean
        Return Not value1 = value2
    End Operator
    Public Overrides Function Equals(obj As Object) As Boolean
        If TypeOf obj Is SCROLLINFO Then
            Return CType(obj, SCROLLINFO) = Me
        Else
            Return False
        End If
    End Function
    Public Overrides Function ToString() As String
        Return Join({"Size=" + CbSize.ToString(InvariantCulture),
                    "Mask=" + FMask.ToString(InvariantCulture),
                    "Min=" + NMin.ToString(InvariantCulture),
                    "Max=" + NMax.ToString(InvariantCulture),
                    "Page=" + NPage.ToString(InvariantCulture),
                    "Pos=" + NPos.ToString(InvariantCulture),
                    "Track=" + NTrackPos.ToString(InvariantCulture)}, ",")
    End Function
End Structure
Public NotInheritable Class NativeMethods
    Public Enum WindowAction
        SWFORCEMINIMIZE = 11
        SWHIDE = 0
        SWMAXIMIZE = 3
        SWMINIMIZE = 6
        SWRESTORE = 9
        SWSHOW = 5
        SWSHOWDEFAULT = 10
        SWSHOWMAXIMIZED = 3
        SWSHOWMINIMIZED = 2
        SWSHOWMINNOACTIVE = 7
        SWSHOWNA = 8
        SWSHOWNOACTIVATE = 4
        SWSHOWNORMAL = 1
    End Enum
    Public Shared Sub WindowShowHide(hwnd As IntPtr, Action As WindowAction)
        Dim actionIndex As Integer = Action
        ShowWindow(hwnd, actionIndex)
    End Sub
    Public Shared Sub WindowMove(hwnd As IntPtr, x As Integer, Y As Integer, Width As Integer, Height As Integer, Repaint As Boolean)
        MoveWindow(hwnd, x, Y, Width, Height, Repaint)
    End Sub
    Public Shared Sub WindowMinimize(hwnd As IntPtr)
        ShowWindow(hwnd, 2)
    End Sub
    Public Shared Sub WindowHide(hwnd As IntPtr)
        ShowWindow(hwnd, 0)
    End Sub

    Private Sub New()
    End Sub
    Friend Declare Function SetProcessDPIAware Lib "user32.dll" () As Boolean
    <DllImport("user32.dll", EntryPoint:="GetScrollInfo")>
    Friend Shared Function GetScrollInfo(hwnd As IntPtr, nBar As Integer, ByRef lpsi As SCROLLINFO) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Friend Shared Function GetScrollPos(hWnd As IntPtr, nBar As Integer) As Integer
    End Function
    <DllImport("user32.dll")>
    Friend Shared Function SetScrollPos(hWnd As IntPtr, nBar As Integer, nPos As Integer, bRedraw As Boolean) As Integer
    End Function
    Friend Declare Function PostMessageA Lib "user32.dll" (hwnd As IntPtr, wMsg As Integer, wParam As Integer, lParam As Integer) As Boolean
    <DllImport("user32.dll")>
    Friend Shared Function GetCursorPos(ByRef lpPoint As Point) As Boolean
    End Function
    <DllImport("user32.dll")>
    Friend Shared Function SetCursorPos(x As Integer, Y As Integer) As Boolean
    End Function
    <DllImport("User32.dll", CharSet:=CharSet.Auto)>
    Friend Shared Function ReleaseDC(hWnd As IntPtr, hDC As IntPtr) As Integer
    End Function
    <DllImport("User32.dll")>
    Friend Shared Function GetWindowDC(hWnd As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Friend Shared Function SendInput(nInputs As UInteger, ByRef pInputs As INPUT, cbSize As Integer) As UInteger
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Friend Shared Function SendInput(numberOfInputs As UInteger, inputs As INPUT(), sizeOfInputStructure As Integer) As UInteger
    End Function
    <DllImport("kernel32.dll", CallingConvention:=CallingConvention.Winapi, SetLastError:=True)>
    Friend Shared Function IsWow64Process(<[In]()> hProcess As IntPtr, <Out()> ByRef wow64Process As Boolean) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Friend Shared Function MessageBeep(uType As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function
    <DllImport("user32.dll", EntryPoint:="CreateIconIndirect")>
    Friend Shared Function CreateIconIndirect(iconInfo As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Friend Shared Function DestroyIcon(handle As IntPtr) As Boolean
    End Function
    <DllImport("gdi32.dll")>
    Friend Shared Function DeleteObject(hObject As IntPtr) As Boolean
    End Function
    <DllImport("gdi32.dll")> Friend Shared Function GetDeviceCaps(hdc As IntPtr, nIndex As Integer) As Integer
    End Function
    <DllImport("user32.dll")> Private Shared Function GetDC(hWnd As IntPtr) As IntPtr
    End Function
    Friend Declare Auto Function GetSystemMetrics Lib "user32.dll" (smIndex As Integer) As Integer
    Friend Declare Function GetKeyState Lib "user32.dll" (nVirtKey As Integer) As Short
    Friend Declare Function SetForegroundWindow Lib "user32.dll" (hwnd As IntPtr) As Integer
    Friend Declare Function GetWindowPlacement Lib "user32" (hwnd As IntPtr, ByRef lpwndpl As WindowPlacement) As Long
    Friend Declare Function SetWindowPlacement Lib "user32" (hwnd As IntPtr, ByRef lpwndpl As WindowPlacement) As Long
    Friend Declare Function GetWindowThreadProcessId Lib "User32" (HWND As Long, lpdwProcessId As Long) As Long
    Friend Declare Function IsIconic Lib "User32" (HWND As Long) As Long
    <DllImport("user32.dll", SetLastError:=True)> Private Shared Function ShowWindow(HWND As IntPtr, nCmdShow As Integer) As Long
    End Function
    Friend Declare Function AttachThreadInput Lib "User32" (idAttach As Long, idAttachTo As Long, fAttach As Long) As Long
    Friend Declare Function GetForegroundWindow Lib "User32" () As Long
    Friend Declare Function GetDesktopWindow Lib "User32" () As Long
    Friend Declare Function GetWindowRect Lib "User32" (HWND As Long, lpRect As RECT) As Long
    Friend Declare Function MoveWindow Lib "User32" (HWND As IntPtr, x As Integer, Y As Integer, nWidth As Integer, nHeight As Integer, bRepaint As Boolean) As Long
    Friend Declare Function EnumWindows Lib "User32" (lpEnumFunc As Long, lParam As Long) As Long
    Friend Declare Function EnumChildWindows Lib "User32" (hWndParent As Long, lpEnumFunc As Long, lParam As Long) As Long
    Friend Declare Function EnumThreadWindows Lib "User32" (dwThreadId As Long, lpfn As Long, lParam As Long) As Long
    Friend Declare Function BringWindowToTop Lib "User32" (HWND As Long) As Long
    Friend Declare Function SetActiveWindow Lib "user32.dll" (HWND As Long) As Long
    Friend Declare Function IsWindowVisible Lib "user32.dll" (HWND As Long) As Boolean
    Friend Declare Function SendMessage Lib "User32" Alias "SendMessageA" (HWND As Long, wMsg As Long, wParam As Long, lParam As Long) As Long
    Friend Declare Sub Mouse_Event Lib "user32.dll" Alias "mouse_event" (dwFlags As Integer, dx As Integer, dy As Integer, cButtons As Integer, dwExtraInfo As Integer)
    <DllImport("user32.dll", EntryPoint:="GetClassLong")> Friend Shared Function GetClassLong(hWnd As IntPtr, nIndex As Integer) As Integer
    End Function
    <DllImport("user32.dll", EntryPoint:="SetClassLong")> Friend Shared Function SetClassLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As Integer) As Integer
    End Function
#Region " S T A R T  /  S T O P   D R A W I N G "
    Private Const WM_SETREDRAW As Integer = &HB
    Private Const WM_USER As Integer = &H400
    Private Const EM_GETEVENTMASK As Integer = WM_USER + 59
    Private Const EM_SETEVENTMASK As Integer = WM_USER + 69
    Private Shared EventMask As IntPtr
    <DllImport("user32", CharSet:=CharSet.Auto)> Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As Integer, lParam As IntPtr) As IntPtr
    End Function
    Public Shared Sub StopDrawing(drawControl As Control)

        If drawControl IsNot Nothing Then
            SendMessage(drawControl.Handle, WM_SETREDRAW, 0, IntPtr.Zero)
            'Stop sending of events
            EventMask = SendMessage(drawControl.Handle, EM_GETEVENTMASK, 0, IntPtr.Zero)
        End If

    End Sub
    Public Shared Sub StartDrawing(drawControl As Control, Optional refresh As Boolean = True)

        If drawControl IsNot Nothing Then
            SendMessage(drawControl.Handle, EM_SETEVENTMASK, 0, EventMask)
            'turn on redrawing
            SendMessage(drawControl.Handle, WM_SETREDRAW, 1, IntPtr.Zero)
            If refresh Then drawControl.Invalidate()
        End If

    End Sub
#End Region
    Friend Structure WindowPlacement
        Dim Length As Integer
        Dim Flags As Integer
        Dim ShowCmd As Integer
        Dim ptMinPosition As POINTAPI
        Dim ptMaxPosition As POINTAPI
        Dim rcNormalPosition As RECT
    End Structure
    Friend Structure POINTAPI
        Dim X As Integer
        Dim Y As Integer
    End Structure
    Friend Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure
End Class
#End Region

Public NotInheritable Class CustomToolTip
    Inherits ToolTip
    Public Enum ShowPosition
        Above
        AboveLeft
        AboveRight
        Below
        BelowLeft
        BelowRight
        Left
        Right
    End Enum
    Public Sub New()

        Active = True
        AutomaticDelay = 100
        InitialDelay = 100
        ShowAlways = False
        StripAmpersands = False
        UseAnimation = True
        UseFading = True

        IsBalloon = False   'Draw Event won't work for IsBalloon=True!
        AutoPopDelay = 5000
        ReshowDelay = 10
        OwnerDraw = True

        _BoundsImage = If(TipImage Is Nothing, Nothing, New Rectangle(2, 2, TipImage.Width, TipImage.Height))

        ' NB Show Sub 1st, then Popup fires, then Draw

    End Sub
    Public Property PreferredPosition As ShowPosition = ShowPosition.AboveLeft
    Private ReadOnly Property TipSize As Size
    Private ReadOnly Property TipText As String
    Private ReadOnly Property BackgroundImage As Image
    Public ReadOnly Property Parent As Control
    Public ReadOnly Property Bounds As Rectangle
        Get
            Return BoundsTip(PreferredPosition)
        End Get
    End Property
    Private ReadOnly Property BoundsParent As Rectangle
    Private ReadOnly Property BoundsText As New SpecialDictionary(Of ShowPosition, Rectangle)
    Private ReadOnly Property BoundsTip As New SpecialDictionary(Of ShowPosition, Rectangle)
    Private ReadOnly Property BoundsImage As Rectangle
    Private ReadOnly Property PointsParent As New SpecialDictionary(Of ShowPosition, Point)
    Private ReadOnly Property PointsConnector As New SpecialDictionary(Of ShowPosition, Point())
    Private ReadOnly Property OffsetShow As Point
    Public ReadOnly Property Words As SpecialDictionary(Of Integer, SpecialDictionary(Of Rectangle, String))
    Public ReadOnly Property WordList As New SpecialDictionary(Of Rectangle, String)
    Public ReadOnly Property DisplayFactor As Single

    Private Font_ As New Font("IBM Plex Mono", 10)
    Public Property TipFont As Font
        Get
            Return Font_
        End Get
        Set(value As Font)
            If value IsNot Nothing Then
                Font_ = value
            End If
        End Set
    End Property
    Public Property TipImage As Image
    Public Property BorderColor As Color = Color.White
    Public Property CaptionColor As Color = Color.WhiteSmoke
    Private TipAlpha As Integer = 255
    Private WithEvents TimerAlpha As New IconTimer With {.Interval = 20, .Counter = 255}
    Public ReadOnly Property Handle As IntPtr

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property
    Private Sub TimerAlpha_Tick() Handles TimerAlpha.Tick

        '10 seconds to fade out with 255
        '1 alpha lower, 255 times, #Seconds to fade / full alpha (255), 5 seconds ==>5,000 ms ==> 5,000 / 255 = 39.2, ie) .Interval=39, Counter=255

        TipAlpha -= 1
        If TipAlpha = 0 Then TimerAlpha.Stop()
        Invalidate(Graphics.FromHwnd(Handle))

    End Sub
    Private Sub Tip_PopUp(sender As Object, e As PopupEventArgs) Handles Me.Popup

        '/// This Sub is called BEFORE Draw. Need to calculate best location and sizes + can only set the e.ToolTipSize here
        _DisplayFactor = DisplayScale()

#Region " KILL THE SHADOW "
        Dim GCL_STYLE As Integer = -26
        Dim CS_DROPSHADOW As Integer = &H20000
        _Handle = CType(GetType(ToolTip).GetProperty("Handle", BindingFlags.NonPublic Or BindingFlags.Instance).GetValue(Me), IntPtr)
        Dim cs = NativeMethods.GetClassLong(Handle, GCL_STYLE)

        If (cs And CS_DROPSHADOW) = CS_DROPSHADOW Then
            cs = cs And Not CS_DROPSHADOW
            Dim result = NativeMethods.SetClassLong(Handle, GCL_STYLE, cs)
        End If
#End Region

        If Not BoundsTip.Any Then Properties_Fill(GetToolTip(e.AssociatedControl), e.AssociatedControl)
        _TipSize = Bounds.Size
        e.ToolTipSize = TipSize

    End Sub
    Private Sub Tip_Draw(sender As Object, e As DrawToolTipEventArgs) Handles Me.Draw
        Invalidate(e.Graphics)
    End Sub
    Private Sub Invalidate(g As Graphics)

        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
#Region " DRAW 'TRANSPARENT' BACKGROUND "
        Dim boundsRelative As New Rectangle(0, 0, Bounds.Width, Bounds.Height)
        g.DrawImage(BackgroundImage, boundsRelative)
        Using penBorder As New Pen(Brushes.Red, 2)
            'g.DrawRectangle(Pens.Red, New Rectangle(0, 0, CInt(boundsRelative.Width - penBorder.Width), CInt(boundsRelative.Height - penBorder.Width)))
        End Using
#End Region
#Region " DRAW CONNECTORS BETWEEN THE TIPTEXT TO THE PARENT CONTROL "
        Using borderBrush As New SolidBrush(Color.FromArgb(TipAlpha, BorderColor))
            Using penConnector As New Pen(borderBrush, 3)
                penConnector.StartCap = Drawing2D.LineCap.RoundAnchor
                penConnector.EndCap = Drawing2D.LineCap.Flat
                Dim pointsConnect = PointsConnector(PreferredPosition)
                Dim pointParent As New Point(pointsConnect.First.X - Bounds.X, pointsConnect.First.Y - Bounds.Y)
                Dim pointTip As Point = New Point(pointsConnect.Last.X - Bounds.X, pointsConnect.Last.Y - Bounds.Y)
                g.DrawLine(penConnector, pointParent, pointTip)
            End Using
        End Using
#End Region
#Region " ROUNDED TEXT SECTION "
        Dim boundsRound As Rectangle = BoundsText(PreferredPosition)
        Using pathBack As Drawing2D.GraphicsPath = DrawRoundedRectangle(boundsRound, 30)
            Using backBrush As New SolidBrush(Color.FromArgb(TipAlpha, BorderColor))
                g.FillPath(backBrush, pathBack)
            End Using
        End Using

        boundsRound.Inflate(-2, -2)
        Using pathBack As Drawing2D.GraphicsPath = DrawRoundedRectangle(boundsRound, 28)
            Using backBrush As New SolidBrush(Color.FromArgb(TipAlpha, CaptionColor))
                g.FillPath(backBrush, pathBack)
            End Using
        End Using
#End Region
#Region " DRAW IMAGE + TEXT "
        boundsRound.Inflate(2, 2)
        Dim wordOffset As New Point(boundsRound.X - boundsRelative.X, boundsRound.Y - boundsRelative.Y)

        If TipImage IsNot Nothing And Words.Any Then
            Dim offsetImageY As Integer = 0
            Dim wordFirstRow = Words.First
            Dim wordFirstBounds = wordFirstRow.Value.First.Key
            Dim rowsPastImage = Words.Where(Function(w) w.Value.Where(Function(r) r.Key.Left < TipImage.Width).Any)
            If rowsPastImage.Any Then
                Dim rowPastImageFirst = rowsPastImage.First
                Dim wordPastImageFirst = rowPastImageFirst.Value.First
                offsetImageY = 2 + CInt((wordPastImageFirst.Key.Top - TipImage.Height) / 2) '4 <== boundsTangle.Inflate(-2, -2) Twice
            End If
            Dim offsetImageX As Integer = 2 + CInt((wordFirstBounds.Left - TipImage.Width) / 2) '4 <== boundsTangle.Inflate(-2, -2) Twice
            Dim boundsImageOffset As New Rectangle(offsetImageX, offsetImageY, TipImage.Width, TipImage.Height)
            boundsImageOffset.Offset(wordOffset)
            g.DrawImage(TipImage, boundsImageOffset)
        End If

        g.TextRenderingHint = Text.TextRenderingHint.AntiAlias
        For Each row As SpecialDictionary(Of Rectangle, String) In Words.Values
            If 0 = 1 Then
                For Each word In row
                    Dim boxWord As Rectangle = word.Key
                    boxWord.Offset(wordOffset)
                    g.DrawRectangle(Pens.Red, boxWord)
                    Using brushFore As New SolidBrush(ForeColor)
                        Using sf As New StringFormat With {
                            .Alignment = StringAlignment.Near,
                            .LineAlignment = StringAlignment.Center,
                            .Trimming = StringTrimming.None
                        }
                            g.DrawString(
        word.Value,
        TipFont,
        brushFore,
        boxWord,
        sf
        )
                        End Using
                    End Using
                Next
            Else
                If row.Any Then
                    Dim wordFirst As Rectangle = row.First.Key
                    Dim rowLeft As Integer = wordFirst.Left
                    Dim rowTop As Integer = wordFirst.Top
                    Dim wordLast As Rectangle = row.Last.Key
                    Dim rowRight As Integer = wordLast.Right
                    Dim rowWidth As Integer = {rowRight - rowLeft, boundsRound.Right}.Max
                    Dim rowHeight As Integer = wordLast.Height
                    Dim boxRow As New Rectangle(rowLeft + wordOffset.X, rowTop + wordOffset.Y, rowWidth, rowHeight)
                    Dim line As String = Join(row.Values.ToArray)
                    Using brushFore As New SolidBrush(ForeColor)
                        Using sf As New StringFormat With {
                            .Alignment = StringAlignment.Near,
                            .LineAlignment = StringAlignment.Center,
                            .Trimming = StringTrimming.None
                        }
                            g.DrawString(
        line,
        TipFont,
        brushFore,
        boxRow,
        sf
        )
                        End Using
                    End Using
                End If
            End If
        Next
#End Region

    End Sub
    Public Shadows Sub Show(textShow As String, windowShow As IWin32Window)

        '/// Regular ToolTip has the below Methods to Show
        '1] Text+Window, 2] Text+Window+Duration, 3] Text+Window+Point, 4] Text+Window+x+y, 5] Text+Window+Point+Duration 6] Text+Window++x+y+Duration

        If windowShow IsNot Nothing Then
            If Not BoundsTip.Any Then
                Properties_Fill(textShow, Control.FromHandle(windowShow.Handle))

                'TimerAlpha.Start()     Flickering too annoying on redraw
            End If
            MyBase.Show(textShow, windowShow, OffsetShow) '<== Point is an OFFSET value to the Associated Control's location - NOT the screen position
        End If

    End Sub
    Private Sub Properties_Fill(textTip As String, parentOfMe As Control)

        _TipText = textTip

        _BoundsImage = If(TipImage Is Nothing, Nothing, New Rectangle(2, 2, TipImage.Width, TipImage.Height))
        Dim boundsWords = WordRectangles(TipText, TipFont, BoundsImage)
        _Words = boundsWords.Value
        For Each row In Words
            For Each word In row.Value
                WordList.Add(word)
            Next
        Next

        Dim sizeOfMe As New Size(boundsWords.Key.Width, boundsWords.Key.Height)
        Dim sizeConnector As New Size(24, 24)
        'Determine the position - then if it makes sense
        'Connector between the Control and the Tip will be a rectangle:     80 x 40 | 40 x 80 | 80 x 80
        _Parent = parentOfMe
        Dim parentTopLeft As Point = Parent.PointToScreen(New Point())
        Dim midX As Integer = CInt(Parent.Width / 2)
        Dim midY As Integer = CInt(Parent.Height / 2)
        EnumNames(GetType(ShowPosition)).ForEach(Sub(sp)
                                                     Dim position As ShowPosition = ParseEnum(Of ShowPosition)(sp)
                                                     Dim isLeft As Boolean = sp.Contains("Left") : Dim isRight As Boolean = sp.Contains("Right")
                                                     Dim isMidH As Boolean = Not (isLeft Or isRight)
                                                     Dim isAbove As Boolean = sp.Contains("Above") : Dim isBelow As Boolean = sp.Contains("Below")
                                                     Dim isMidV As Boolean = Not (isAbove Or isBelow)
                                                     Dim xPos As Integer = parentTopLeft.X + If(isLeft, 0, If(isRight, Parent.Width, CInt(Parent.Width / 2)))
                                                     Dim yPos As Integer = parentTopLeft.Y + If(isAbove, 0, If(isBelow, Parent.Height, CInt(Parent.Height / 2)))
                                                     'PointsParent is 1st of 2 connector points so a line can be drawn between the TipCorner and the Parent
                                                     Dim pointParent As New Point(xPos, yPos)
                                                     PointsParent.Add(position, pointParent)
                                                     'Only offset X if Left or Right, Y if Above or Below
                                                     Dim xOff As Integer = If(isLeft, -sizeConnector.Width, If(isRight, sizeConnector.Width, 0))
                                                     Dim yOff As Integer = If(isAbove, -sizeConnector.Height, If(isBelow, sizeConnector.Height, 0))
                                                     'PointsConnector is the 2nd of 2 connector points between the TipCorner and the Parent
                                                     Dim pointTip As New Point(pointParent.X + xOff, pointParent.Y + yOff)
                                                     Dim pointX As Integer = If(isLeft, pointTip.X - sizeOfMe.Width, If(isRight, pointParent.X, pointParent.X - CInt(sizeOfMe.Width / 2)))
                                                     Dim pointY As Integer = If(isAbove, pointTip.Y - sizeOfMe.Height, If(isBelow, pointParent.Y, pointParent.Y - CInt(sizeOfMe.Height / 2)))
                                                     Dim sizeW As Integer = sizeOfMe.Width + If(isMidH, 0, sizeConnector.Width)
                                                     Dim sizeH As Integer = sizeOfMe.Height + If(isMidV, 0, sizeConnector.Height)
                                                     BoundsTip.Add(position, New Rectangle(pointX, pointY, sizeW, sizeH))
                                                     Dim pointXX As Integer = If(isRight, pointTip.X, pointX)
                                                     Dim pointYY As Integer = If(isBelow, pointTip.Y, pointY)
                                                     BoundsText.Add(position, New Rectangle(pointXX - Bounds.X, pointYY - Bounds.Y, sizeOfMe.Width, sizeOfMe.Height))
                                                     'Below moves the connector line between the Tip and the Control in toward the Tip/away from the Control which allows space for the EndCap to draw at the Control and connects the line to the rounded Tip box
                                                     Dim lineX As Integer = If(isLeft, -3, If(isRight, 3, 0)) 'Assumes Connector is at 45° since using sizeConnector(x, y) where x=y
                                                     Dim lineY As Integer = If(isAbove, -3, If(isBelow, 3, 0)) 'Assumes Connector is at 45° since using sizeConnector(x, y) where x=y
                                                     Dim lineOffset As New Point(lineX, lineY)
                                                     Dim pointP As Point = pointParent
                                                     pointP.Offset(lineOffset)
                                                     Dim pointT As Point = pointTip
                                                     pointT.Offset(lineOffset)
                                                     pointT.Offset(lineOffset) 'Twice is better
                                                     PointsConnector.Add(position, {pointP, pointT})
                                                     'If position = PreferredPosition Then Stop
                                                 End Sub)
        _BoundsParent = New Rectangle(PointsParent(ShowPosition.AboveLeft), Parent.Size) 'AboveLeft is the upper left corner of the Associated Control
        _OffsetShow = New Point(Bounds.X - BoundsParent.X, Bounds.Y - BoundsParent.Y)
        '┌───┐       ┌───┐       ┌───┐
        '│A.L│       │ABV│       │A.R│
        '└───┘┐      └─┬─┘      ┌└───┘
        '      \       │       /
        '┌───┐  ┌─────────────┐  ┌───┐
        '│LFT├──┤ASSOC.CONTROL├──┤RGT│
        '└───┘  └──────┬──────┘  └───┘
        '      /       │       \
        '┌───┐┘      ┌─┴─┐      └┌───┐
        '│B.L│       │BLW│       │B.R│
        '└───┘       └───┘       └───┘
        _DisplayFactor = DisplayScale()
        Dim boundsScale As Rectangle = Bounds
        Dim bmpCapture As Bitmap = New Bitmap(CInt(boundsScale.Width * DisplayFactor), CInt(boundsScale.Height * DisplayFactor))
        Using g As Graphics = Graphics.FromImage(bmpCapture)
            g.CopyFromScreen(CInt(boundsScale.X * DisplayFactor),
                 CInt(boundsScale.Y * DisplayFactor),
                 0,
                 0,
                 bmpCapture.Size,
                 CopyPixelOperation.SourceCopy)
        End Using
        _BackgroundImage = bmpCapture

    End Sub
End Class

Public NotInheritable Class ClipboardHelper
    Private Sub New()
    End Sub
#Region " Fields and Constants "

    ''' <summary>      
    ''' The string contains index references to  other spots in the string, so we need placeholders so we can compute the  offsets. <br/>      
    ''' The  <![CDATA[<<<<<<<]]>_ strings are just placeholders.  We'll back-patch them actual values afterwards. <br/>      
    ''' The string layout  (<![CDATA[<<<]]>) also ensures that it can't appear in the body  of the html because the <![CDATA[<]]> <br/>      
    ''' character must be escaped. <br/>      
    ''' </summary>      
    Private Const Header As String = "Version:0.9      " & vbCr & vbLf & "StartHTML:<<<<<<<<1      " & vbCr & vbLf & "EndHTML:<<<<<<<<2      " & vbCr & vbLf & "StartFragment:<<<<<<<<3      " & vbCr & vbLf & "EndFragment:<<<<<<<<4      " & vbCr & vbLf & "StartSelection:<<<<<<<<3      " & vbCr & vbLf & "EndSelection:<<<<<<<<4"

    ''' <summary>      
    ''' html comment to point the beginning of  html fragment      
    ''' </summary>      
    Public Const StartFragment As String = "<!--StartFragment-->"

    ''' <summary>      
    ''' html comment to point the end of html  fragment      
    ''' </summary>      
    Public Const EndFragment As String = "<!--EndFragment-->"

    ''' <summary>      
    ''' Used to calculate characters byte count  in UTF-8      
    ''' </summary>      
    Private Shared ReadOnly _byteCount As Char() = New Char(0) {}

#End Region
    ''' <summary>      
    ''' Create <see  cref="DataObject"/> with given html and plain-text ready to be  used for clipboard or drag and drop.<br/>      
    ''' Handle missing  <![CDATA[<html>]]> tags, specified startend segments and Unicode  characters.      
    ''' </summary>      
    ''' <remarks>      
    ''' <para>      
    ''' Windows Clipboard works with UTF-8  Unicode encoding while .NET strings use with UTF-16 so for clipboard to  correctly      
    ''' decode Unicode string added to it from  .NET we needs to be re-encoded it using UTF-8 encoding.      
    ''' </para>      
    ''' <para>      
    ''' Builds the CF_HTML header correctly for  all possible HTMLs<br/>      
    ''' If given html contains start/end  fragments then it will use them in the header:      
    '''  <code><![CDATA[<html><body><!--StartFragment-->hello  <b>world</b><!--EndFragment--></body></html>]]></code>      
    ''' If given html contains html/body tags  then it will inject start/end fragments to exclude html/body tags:      
    '''  <code><![CDATA[<html><body>hello  <b>world</b></body></html>]]></code>      
    ''' If given html doesn't contain html/body  tags then it will inject the tags and start/end fragments properly:      
    ''' <code><![CDATA[hello  <b>world</b>]]></code>      
    ''' In all cases creating a proper CF_HTML  header:<br/>      
    ''' <code>      
    ''' <![CDATA[      
    ''' Version:1.0      
    ''' StartHTML:000000177      
    ''' EndHTML:000000329      
    ''' StartFragment:000000277      
    ''' EndFragment:000000295      
    ''' StartSelection:000000277      
    ''' EndSelection:000000277      
    ''' <!DOCTYPE HTML PUBLIC  "-//W3C//DTD HTML 4.0 Transitional//EN">      
    '''  <html><body><!--StartFragment-->hello  <b>world</b><!--EndFragment--></body></html>      
    ''' ]]>      
    ''' </code>      
    ''' See format specification here: [http://msdn.microsoft.com/library/default.asp?url=/workshop/networking/clipboard/htmlclipboard.asp][9]      
    ''' </para>      
    ''' </remarks>      
    ''' <param name="html">a  html fragment</param>      
    ''' <param  name="plainText">the plain text</param>      
    Public Shared Function CreateDataObject(html As String, plainText As String) As DataObject

        html = If(html, String.Empty)
        Dim htmlFragment = GetHtmlDataString(html)
        ' re-encode the string so it will work  correctly (fixed in CLR 4.0)      
        If Environment.Version.Major < 4 AndAlso html.Length <> Encoding.UTF8.GetByteCount(html) Then
            htmlFragment = Encoding.[Default].GetString(Encoding.UTF8.GetBytes(htmlFragment))
        End If

        Dim dataObject = New DataObject()
        dataObject.SetData(DataFormats.Html, htmlFragment)
        dataObject.SetData(DataFormats.Text, plainText)
        dataObject.SetData(DataFormats.UnicodeText, plainText)
        Return dataObject

    End Function

    ''' <summary>      
    ''' Clears clipboard and sets the given  HTML and plain text fragment to the clipboard, providing additional  meta-information for HTML.<br/>      
    ''' See <see  cref="CreateDataObject"/> for HTML fragment details.<br/>      
    ''' </summary>      
    ''' <example>      
    '''  ClipboardHelper.CopyToClipboard("Hello <b>World</b>",  "Hello World");      
    ''' </example>      
    ''' <param name="html">a  html fragment</param>      
    Public Shared Sub CopyToClipboard(html As String)
        Dim dataObject = CreateDataObject(html, html)
        Clipboard.SetDataObject(dataObject, True)
    End Sub
    Public Shared Sub CopyToClipboard(html As String, plainText As String)
        Dim dataObject = CreateDataObject(html, plainText)
        Clipboard.SetDataObject(dataObject, True)
    End Sub

    ''' <summary>      
    ''' Generate HTML fragment data string with  header that is required for the clipboard.      
    ''' </summary>      
    ''' <param name="html">the  html to generate for</param>      
    ''' <returns>the resulted  string</returns>      
    Private Shared Function GetHtmlDataString(html As String) As String
        Dim sb = New StringBuilder()
        sb.AppendLine(Header)
        sb.AppendLine("<!DOCTYPE HTML  PUBLIC ""-//W3C//DTD HTML 4.0  Transitional//EN"">")

        ' if given html already provided the  fragments we won't add them      
        Dim fragmentStart As Integer, fragmentEnd As Integer
        Dim fragmentStartIdx As Integer = html.IndexOf(StartFragment, StringComparison.OrdinalIgnoreCase)
        Dim fragmentEndIdx As Integer = html.LastIndexOf(EndFragment, StringComparison.OrdinalIgnoreCase)

        ' if html tag is missing add it  surrounding the given html (critical)      
        Dim htmlOpenIdx As Integer = html.IndexOf("<html", StringComparison.OrdinalIgnoreCase)
        Dim htmlOpenEndIdx As Integer = If(htmlOpenIdx > -1, html.IndexOf(">"c, htmlOpenIdx) + 1, -1)
        Dim htmlCloseIdx As Integer = html.LastIndexOf("</html", StringComparison.OrdinalIgnoreCase)

        If fragmentStartIdx < 0 AndAlso fragmentEndIdx < 0 Then
            Dim bodyOpenIdx As Integer = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase)
            Dim bodyOpenEndIdx As Integer = If(bodyOpenIdx > -1, html.IndexOf(">"c, bodyOpenIdx) + 1, -1)

            If htmlOpenEndIdx < 0 AndAlso bodyOpenEndIdx < 0 Then
                ' the given html doesn't  contain html or body tags so we need to add them and place start/end fragments  around the given html only      
                sb.Append("<html><body>")
                sb.Append(StartFragment)
                fragmentStart = GetByteCount(sb)
                sb.Append(html)
                fragmentEnd = GetByteCount(sb)
                sb.Append(EndFragment)
                sb.Append("</body></html>")
            Else
                ' insert start/end fragments  in the proper place (related to html/body tags if exists) so the paste will  work correctly      
                Dim bodyCloseIdx As Integer = html.LastIndexOf("</body", StringComparison.OrdinalIgnoreCase)

                If htmlOpenEndIdx < 0 Then
                    sb.Append("<html>")
                Else
                    sb.Append(html, 0, htmlOpenEndIdx)
                End If

                If bodyOpenEndIdx > -1 Then
                    sb.Append(html, If(htmlOpenEndIdx > -1, htmlOpenEndIdx, 0), bodyOpenEndIdx - (If(htmlOpenEndIdx > -1, htmlOpenEndIdx, 0)))
                End If

                sb.Append(StartFragment)
                fragmentStart = GetByteCount(sb)

                Dim innerHtmlStart = If(bodyOpenEndIdx > -1, bodyOpenEndIdx, (If(htmlOpenEndIdx > -1, htmlOpenEndIdx, 0)))
                Dim innerHtmlEnd = If(bodyCloseIdx > -1, bodyCloseIdx, (If(htmlCloseIdx > -1, htmlCloseIdx, html.Length)))
                sb.Append(html, innerHtmlStart, innerHtmlEnd - innerHtmlStart)

                fragmentEnd = GetByteCount(sb)
                sb.Append(EndFragment)

                If innerHtmlEnd < html.Length Then
                    sb.Append(html, innerHtmlEnd, html.Length - innerHtmlEnd)
                End If

                If htmlCloseIdx < 0 Then
                    sb.Append("</html>")
                End If
            End If
        Else
            ' handle html with existing  startend fragments just need to calculate the correct bytes offset (surround  with html tag if missing)      
            If htmlOpenEndIdx < 0 Then
                sb.Append("<html>")
            End If
            Dim start As Integer = GetByteCount(sb)
            sb.Append(html)
            fragmentStart = start + GetByteCount(sb, start, start + fragmentStartIdx) + StartFragment.Length
            fragmentEnd = start + GetByteCount(sb, start, start + fragmentEndIdx)
            If htmlCloseIdx < 0 Then
                sb.Append("</html>")
            End If
        End If

        ' Back-patch offsets (scan only the  header part for performance)      
        sb.Replace("<<<<<<<<1", Header.Length.ToString("D9", InvariantCulture), 0, Header.Length)
        sb.Replace("<<<<<<<<2", GetByteCount(sb).ToString("D9", InvariantCulture), 0, Header.Length)
        sb.Replace("<<<<<<<<3", fragmentStart.ToString("D9", InvariantCulture), 0, Header.Length)
        sb.Replace("<<<<<<<<4", fragmentEnd.ToString("D9", InvariantCulture), 0, Header.Length)

        Return sb.ToString()
    End Function

    ''' <summary>      
    ''' Calculates the number of bytes produced  by encoding the string in the string builder in UTF-8 and not .NET default  string encoding.      
    ''' </summary>      
    ''' <param name="sb">the  string builder to count its string</param>      
    ''' <param  name="start">optional: the start index to calculate from (default  - start of string)</param>      
    ''' <param  name="end">optional: the end index to calculate to (default - end  of string)</param>      
    ''' <returns>the number of bytes  required to encode the string in UTF-8</returns>      
    Private Shared Function GetByteCount(sb As StringBuilder, Optional start As Integer = 0, Optional [end] As Integer = -1) As Integer
        Dim count As Integer = 0
        [end] = If([end] > -1, [end], sb.Length)
        For i As Integer = start To [end] - 1
            _byteCount(0) = sb(i)
            count += Encoding.UTF8.GetByteCount(_byteCount)
        Next
        Return count
    End Function
End Class

Public NotInheritable Class DictionaryEventArgs
    Inherits EventArgs
    Public ReadOnly Property Key As Object
    Public ReadOnly Property Value As Object
    Public ReadOnly Property LastValue As Object
    Public Sub New(key As Object, oldValue As Object, newValue As Object)

        Me.Key = key
        LastValue = oldValue
        Value = newValue

    End Sub
End Class
<Serializable>
Public NotInheritable Class SpecialDictionary(Of TKey, TValue)
    Inherits Dictionary(Of TKey, TValue)
    Implements IDictionary(Of TKey, TValue)
    Public Event PropertyChanged(sender As Object, e As DictionaryEventArgs)
    Public Sub New()
        MyBase.New()
    End Sub
    Public Sub New(kvpEnumerable As IEnumerable(Of KeyValuePair(Of TKey, TValue)))
        AddRange(kvpEnumerable)
    End Sub
    Public Sub New(kvpList As List(Of KeyValuePair(Of TKey, TValue)))
        AddRange(kvpList)
    End Sub
    Public Sub New(dict As Dictionary(Of TKey, TValue))
        AddRange(dict)
    End Sub
    Public Sub New(dict As SpecialDictionary(Of TKey, TValue))
        AddRange(dict)
    End Sub
    Private Sub New(serializationInfo As Runtime.Serialization.SerializationInfo, streamingContext As Runtime.Serialization.StreamingContext)
        Throw New NotImplementedException()
    End Sub
    Public Property Tag As Object
    Private KeyExists As TriState
    Public Overloads Function Add(kvp As KeyValuePair(Of TKey, TValue)) As KeyValuePair(Of TKey, TValue)
        Return Add(kvp.Key, kvp.Value)
    End Function
    Public Overloads Function Add(key As TKey, value As TValue) As KeyValuePair(Of TKey, TValue)
        MyBase.Add(key, value)
        Return New KeyValuePair(Of TKey, TValue)(key, value)
    End Function
    Public Sub AddRange(dict As Dictionary(Of TKey, TValue))

        If dict IsNot Nothing Then
            For Each kvp In dict
                Add(kvp.Key, kvp.Value)
            Next
        End If

    End Sub
    Public Sub AddRange(dict As SpecialDictionary(Of TKey, TValue))

        If dict IsNot Nothing Then
            For Each kvp In dict
                Add(kvp.Key, kvp.Value)
            Next
        End If

    End Sub
    Public Sub AddRange(kvpEnumerable As IEnumerable(Of KeyValuePair(Of TKey, TValue)))

        If kvpEnumerable IsNot Nothing Then
            For Each kvp In kvpEnumerable
                Add(kvp.Key, kvp.Value)
            Next
        End If

    End Sub
    Public Sub AddRange(kvpList As List(Of KeyValuePair(Of TKey, TValue)))

        If kvpList IsNot Nothing Then
            kvpList.ForEach(Sub(kvp)
                                Add(kvp.Key, kvp.Value)
                            End Sub)
        End If

    End Sub
    Public Shadows Function ContainsKey(key As TKey) As Boolean
        Dim value = Item(key)
        Return KeyExists = TriState.True
    End Function
    Default Public Overloads Property Item(key As TKey) As TValue
        Get
            KeyExists = TriState.UseDefault
            If key.GetType Is GetType(String) Then
                Dim matchingKey As TKey = Nothing
                For Each keyItem As TKey In Keys
                    If keyItem.ToString.ToUpperInvariant = key.ToString.ToUpperInvariant Then
                        matchingKey = keyItem
                        Exit For
                    End If
                Next
                If matchingKey Is Nothing Then
                    KeyExists = TriState.False
                    Return Nothing
                Else
                    KeyExists = TriState.True
                    Return MyBase.Item(matchingKey)
                End If
            Else
                If MyBase.ContainsKey(key) Then
                    KeyExists = TriState.True
                    Return MyBase.Item(key)
                Else
                    KeyExists = TriState.False
                    Return Nothing
                End If
            End If
        End Get
        Set(value As TValue)
            RaiseEvent PropertyChanged(Me, New DictionaryEventArgs(key, Me(key), value))
            MyBase.Item(key) = value
        End Set
    End Property
    Public Sub SortByKeys(Optional ascending As Boolean = True)

        Dim kvpList As New List(Of KeyValuePair(Of TKey, TValue))
        For Each kvp In Me
            kvpList.Add(New KeyValuePair(Of TKey, TValue)(kvp.Key, kvp.Value))
        Next
        If ascending Then
            kvpList.Sort(Function(x, y)
                             Select Case GetType(TKey)
                                 Case GetType(Byte), GetType(Short), GetType(Integer), GetType(Long)
                                     Return CLng(x.Key.ToString).CompareTo(CLng(y.Key.ToString))

                                 Case GetType(Boolean)
                                     Return CBool(x.Key.ToString).CompareTo(CBool(y.Key.ToString))

                                 Case GetType(Date)
                                     Return CDate(x.Key.ToString).CompareTo(CDate(y.Key.ToString))

                                 Case GetType(String)
                                     Return String.Compare(x.Key.ToString, y.Key.ToString, True, InvariantCulture)

                                 Case Else
                                     Return String.Compare(x.Key.ToString, y.Key.ToString, True, InvariantCulture)

                             End Select
                         End Function)
        Else
            kvpList.Sort(Function(y, x)
                             Select Case GetType(TKey)
                                 Case GetType(Byte), GetType(Short), GetType(Integer), GetType(Long)
                                     Return CLng(x.Key.ToString).CompareTo(CLng(y.Key.ToString))

                                 Case GetType(Boolean)
                                     Return CBool(x.Key.ToString).CompareTo(CBool(y.Key.ToString))

                                 Case GetType(Date)
                                     Return CDate(x.Key.ToString).CompareTo(CDate(y.Key.ToString))

                                 Case GetType(String)
                                     Return String.Compare(x.Key.ToString, y.Key.ToString, True, InvariantCulture)

                                 Case Else
                                     Return String.Compare(x.Key.ToString, y.Key.ToString, True, InvariantCulture)

                             End Select
                         End Function)
        End If
        Clear()
        kvpList.ForEach(Sub(kvp)
                            Add(kvp.Key, kvp.Value)
                        End Sub)

    End Sub
    Public Sub SortByValues(Optional ascending As Boolean = True)

        Dim kvpList As New List(Of KeyValuePair(Of TKey, TValue))
        For Each kvp In Me
            kvpList.Add(New KeyValuePair(Of TKey, TValue)(kvp.Key, kvp.Value))
        Next
        If ascending Then
            kvpList.Sort(Function(x, y)
                             Select Case GetType(TValue)
                                 Case GetType(Byte), GetType(Short), GetType(Integer), GetType(Long)
                                     Return CLng(x.Value.ToString).CompareTo(CLng(y.Value.ToString))

                                 Case GetType(Boolean)
                                     Return CBool(x.Value.ToString).CompareTo(CBool(y.Value.ToString))

                                 Case GetType(Date)
                                     Return CDate(x.Value.ToString).CompareTo(CDate(y.Value.ToString))

                                 Case GetType(String)
                                     Return String.Compare(x.Value.ToString, y.Value.ToString, True, InvariantCulture)

                                 Case Else
                                     Return String.Compare(x.Value.ToString, y.Value.ToString, True, InvariantCulture)

                             End Select
                         End Function)
        Else
            kvpList.Sort(Function(y, x)
                             Select Case GetType(TValue)
                                 Case GetType(Byte), GetType(Short), GetType(Integer), GetType(Long)
                                     Return CLng(x.Value.ToString).CompareTo(CLng(y.Value.ToString))

                                 Case GetType(Boolean)
                                     Return CBool(x.Value.ToString).CompareTo(CBool(y.Value.ToString))

                                 Case GetType(Date)
                                     Return CDate(x.Value.ToString).CompareTo(CDate(y.Value.ToString))

                                 Case GetType(String)
                                     Return String.Compare(x.Value.ToString, y.Value.ToString, True, InvariantCulture)

                                 Case Else
                                     Return String.Compare(x.Value.ToString, y.Value.ToString, True, InvariantCulture)

                             End Select
                         End Function)
        End If
        Clear()
        kvpList.ForEach(Sub(kvp)
                            Add(kvp.Key, kvp.Value)
                        End Sub)

    End Sub
    Public ReadOnly Property Stringify As String
        Get
            Dim keysValues As New List(Of String)
            For Each kvp In Me
                Dim stringKey As String = If(kvp.Key Is Nothing, String.Empty, kvp.Key.ToString)
                Dim stringValue As String = If(kvp.Value Is Nothing, String.Empty, kvp.Value.ToString)
                keysValues.Add($"Key={stringKey}, Value={stringValue}")
            Next
            Return Microsoft.VisualBasic.Join(keysValues.ToArray, vbNewLine)
        End Get
    End Property
    Public Overrides Function ToString() As String

        Dim keysValues As New List(Of String)
        For Each kvp In Me
            Dim stringKey As String = If(kvp.Key Is Nothing, String.Empty, kvp.Key.ToString)
            Dim stringValue As String = If(kvp.Value Is Nothing, String.Empty, kvp.Value.ToString)
            keysValues.Add($"Key={stringKey}, Value={stringValue}")
        Next
        Return Microsoft.VisualBasic.Join(keysValues.ToArray, vbNewLine)

    End Function
End Class