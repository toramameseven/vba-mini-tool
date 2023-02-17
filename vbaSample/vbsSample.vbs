dim vbsSample

Class SampleClass1

    dim class1_val
    Public common_val

    function SampleClass1_Func1()
        dim class1_Func1_val
        class1_val = 33
    end function

    private function common()
    end function

    sub log()
        WScript.Echo "this is SampleClass1"
    end sub

End Class


Class SampleClass2

    dim class2_val
    private common_val

    function SampleClass2_Func1()
        dim class2_Func1_val
        class2_val = 1
    end function

    public function common()
    end function

    sub log()
        WScript.Echo "this is SampleClass2"
    end sub
End Class

Function myFunction(val1)

    Dim myFunction_Val
    myFunction_Val = val1

    myFunction =  myFunction_Val

End Function

private Sub Main()
    Dim r
    r = myFunction("return value")
    WScript.Echo r

    Dim c
    set c = new SampleClass1
    c.common_val = 100
    call c.log

    Dim c2
    set c2 = new SampleClass2
    c2.common
    call c2.log

End Sub

Call Main



