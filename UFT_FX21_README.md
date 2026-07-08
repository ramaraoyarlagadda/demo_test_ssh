# UFT FX21 Executable Script

Cleaned UFT/VBScript for FX21 deal creation/amendment validation (FM33M, FX41, FX78).

## Files

| File | Purpose |
|------|---------|
| `FX21_Deal_Validation_Executable.vbs` | Full cleaned executable script |
| `FX21_From_Line60_Snippet.vbs` | Minimal FX21 `Select Case` format to paste from ~line 60 |

## How to run from line 60

1. Keep setup code (FileSystem, DataTable, report files, screenshots).
2. Set `RUN_FROM_FX21_SECTION = True` in the full script, **or** paste the snippet in place of the broken FX21 block.
3. Comment out `On Error Resume Next` while debugging.
4. Fix `keyDealUsingFX21` call based on how it is defined (see below).

## Correct FX21 block format

```vbscript
channelSource = "FX21"
TestStep = 1

Select Case channelSource
    Case "FX21"
        TestStep = TestStep + 1
        Call keyDealUsingFX21(TestStep)   ' Sub - no return value
        ' FX21 logic continues here...
End Select
```

## Closure order (full script)

```vbscript
End Select  ' channelSource
End If      ' ProcessThisRow
Next        ' For RowNum
```

## `Type mismatch: 'keyDealUsingFX21'` — fixed

`keyDealUsingFX21` is a **Sub**, not a Function. It has no return value.

| Wrong (causes Type mismatch) | Correct |
|---|---|
| `dealKeyResult = keyDealUsingFX21(TestStep)` | `Call keyDealUsingFX21(TestStep)` |
| `If keyDealUsingFX21(TestStep) Then` | `Call keyDealUsingFX21(TestStep)` then continue |
