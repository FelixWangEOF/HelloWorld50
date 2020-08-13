.method public static void Show()
{
    .entrypoint
    .maxstack 1

    ldstr "Hello World"
    call void [mscorlib]System.Console::WriteLine(string)
    call valuetype [mscorlib]System.ConsoleKeyInfo [mscorlib]System.Console::ReadKey()
    ret
}