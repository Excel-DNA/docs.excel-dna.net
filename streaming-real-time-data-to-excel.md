---
layout: page
title: "Streaming real-time data to Excel"
---
Gert-Jan van der Kamp has posted a very nice end-to-end example on [streaming real-time data to excel](streaming-real-time-data-to-excel), showing how to create a WCF service and Excel-DNA add-in to stream real-time data into Excel.

The example uses the Reactive Extensions support in Excel-DNA v. 0.30 to push the data to an Excel UDF (using Excel's RTD mechanism behind the scenes), together with a Duplex WCF service providing the data.

There was also this CodePlex discussion about the Excel ThrottleInterval option, which trades off the real-time update frequency against stability of the Excel calculation:

- Hi all, I'm trying to make an example of streaming real time data to Excel over WCF duplex. I have it working, but it only seems to update the screen every 2 odd seconds, even when I send out a datapoint every 0.5 seconds. 
Is there some refresh rate to RTD? I've looked everywhere but can't seem to find this.
Here's the relevant bit where I expose the WCF calls as events in Rx, maybe I'm doing something else wrong?
```csharp
// event boilerplate stuff\n        public delegate void ValueSentHandler(ValueSentEventArgs args);
        public static event ValueSentHandler OnValueSent;
        public class ValueSentEventArgs : EventArgs {
		public double Value { get; private set; }
		public ValueSentEventArgs(double Value) {
			this.Value = Value;
		}
        }
        
        // this is an OperationContract called by WCF
        public void SendValue(double x) {
		// invert method call from WCF into event for Rx
			if (OnValueSent != null) OnValueSent(new ValueSentEventArgs(x));
        }

        [ExcelFunction("Gets the latest value")]
        public static object GetValues() {
			return RxExcel.Observe("GetValues", null,
				() => Observable.Create<double>(
					observer => {
						OnValueSent += d => observer.OnNext(d.Value);
					return Disposable.Empty;
					}));
        }
```

Thanks in advance (And for the awesome Excel-DNA),
Gert-Jan

- Hi Gert-Jan, The Excel RTD FAQ has some information about this: [http://msdn.microsoft.com/en-us/library/office/aa140060(v=office.10).aspx](http://msdn.microsoft.com/en-us/library/office/aa140060(v=office.10).aspx)
Excel has a ThrottleInterval setting global to the Excel instance. This limits the rate at which Excel will retrieve RTD data and recalculate. The default value for this setting in Excel is 2 seconds.  
The value can be set in a running instance by setting Application.RTD.ThrottleInterval, and the default can be changed in the registry.  
If you make this value too small and have some RTD server that updates too often, Excel will recalculate constantly and become unstable, so you do need to take some care.  
Also note that the setting applies to the whole Excel instance, not to a particular RTD server, and is persistent across sessions.  
So if you have a Bloomberg add-in or something, that would be affected too. Excel-DNA itself does not interfere with or set the ThrottleInterval at all.  
There is a small sample in the Excel-DNA distribution that shows how to set and reset the ThrottleInterval.  
See Distribution\\RTD\\RealTimeManagerCS.dna (though the COM code here can by simpler with .NET 4 and the 'dynamic' type or an interop assembly reference).  
Like in that example, I suggest you reset it on startup, and have a manual option for the user to decrease the interval.  
Also note that, even with a low or 0 ThrottleInterval, the RTD interface does not guarantee that every value passed back will appear on the sheet.  
There are other situations (like when the user is editing a formula) where the update is skipped, and only the latest value will eventually appear.  
Regards, Govert

