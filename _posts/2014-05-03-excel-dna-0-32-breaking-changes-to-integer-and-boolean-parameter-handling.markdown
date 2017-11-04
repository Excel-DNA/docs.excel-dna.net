---
layout: post
title: "Excel-DNA 0.32 - Breaking changes to integer and boolean parameter handling"
date: 2014-05-03 07:08:00 -0000
permalink: /2014/05/03/excel-dna-0-32-breaking-changes-to-integer-and-boolean-parameter-handling/
categories: uncategorized, 0.32, conversions
---
Excel-DNA version 0.32 introduces some changes in the parameter conversions applied to integer and boolean parameters. These changes improve compatibility with VBA, and make it easier to provide a consistent implementation when the conversion needs to be explicitly implemented, as for some generated methods.

In Excel-DNA versions before 0.32, UDF functions taking integer and boolean parameters were registered with the C API using the respective types, and hence the conversions were performed by Excel before calling the UDF. In Excel-DNA 0.32, these conversions are performed by Excel-DNA, with the changes discussed here. Affected functions would previously have behaved consistent with `.xll` add-ins made with C/C++, where registered with integer or boolean parameter types.

The new behaviour for integer conversions is that double values passed from Excel to integer parameters in UDFs are converted using the "Round-To-Even" midpoint rounding convention. Previously, positive midpoint values (like `2.5`) were rounded up (to `3`), while negative midpoint values were rounded down (`-2.5` to `-3`), with the exception that `-0.5` was rounded to `0`. `Int64` (`long`) parameters are now also handled consistently.

One exception to the VBA compatibility guideline is that incoming boolean `true` values passed to integer parameters are converted to `1`, rather than `-1` as would be the case with VBA. For this case I consider it more important to be consistent with .NET conventions, whereby boolean `true` values are represented by `1`.

For conversions to boolean parameters, the main change is in how fractional values are converted to booleans. The new version is consistent with VBA - any non-zero value is converted to `true`.

I hope you will agree that the improved consistency is worth making these breaking changes, and that the decision will not cause any unexpected problems. As always, I appreciate any feedback, either directly or via the [Excel-DNA Google group][excel-dna-group].

---

The following snapshot gives a good summary of the changes:

![Param Conversion Table Changes][conversion-table-img]

The functions used are as follows:

{% highlight csharp %}
public static object dnaConvertInt32(int value)
{
    return value;
}
{% endhighlight %}

{% highlight vb %}
Function VbaConvertInteger(value As Integer)
    VbaConvertInteger = value
End Function
{% endhighlight %}

{% highlight csharp %}
public static object dnaConvertInt64(long value)
{
    return value;
}
{% endhighlight %}

{% highlight csharp %}
public static object dnaConvertBoolean(bool value)
{
    return value;
}
{% endhighlight %}

{% highlight vb %}
Function VbaConvertBoolean(value As Boolean)
    VbaConvertBoolean = value
End Function
{% endhighlight %}

[excel-dna-group]: http://groups.google.com/group/exceldna
[conversion-table-img]: /images/param-conversion-changes-v0-32.png "Param Conversion Table Changes"
