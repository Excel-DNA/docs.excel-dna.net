---
layout: page
title: "FSharp Type Inference"
---
When creating UDFs with F#, the flexible type inference might lead to function signatures that are not supported by Excel-DNA, or lead to unexpected results.

```fsharp
let MakeTwo x = 2
```

This doesn't work (the UDF doesn't get registered) since the inferred type is _'a -> int_, so is generic over the argument. This is equivalent to the C# signature:

```fsharp
public int MakeTwo<T>(T input) = { return 2; }
```
However, the following, with explicit typing,  does work:

```fsharp
let MakeTwo (x : float) = 2
```

This would apply to any function that is generic over its input. Another example is:

```fsharp
let AddString x y = x.ToString() + y.ToString()
```

which is of the type a' -> b' -> string and doesn't get exposed as an UDF either.

Adding explicit types removes the generic parameters:

```fsharp
let AddString (x:obj) (y:obj) = x.ToString() + y.ToString()
```

Even the simple example in the distribution can be a concern:

```fsharp
let Add x y = x + y
```
F# infers this function to be of the type int -> int -> int, and if called in Excel as =Add(2.5,3.5) then this function will return 7 not 6.

```fsharp
let Add (x:float) (y:float) = x + y
```
