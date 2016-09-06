using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Management.Automation;
using System.Collections.ObjectModel;

namespace PowerShellExecutionSample
{
    class Program
    {
        static void Main(string[] args)
        {
            PowerShellExecutor t = new PowerShellExecutor();
            
            // scenario 1 (synchronous execution)
            t.ExecuteSynchronously();

            // scenario 2 (asynchronous execution)
            //t.ExecuteAsynchronously();

            // scenario 3 (synchronous execution with the namespace test)
            //t.ExecuteSynchronouslyNamespaceTest();
        }
    }

    /// <summary>
    /// Test class object to instantiate from inside PowerShell script.
    /// </summary>
    public class TestObject
    {
        /// <summary>
        /// Gets or sets the Name property
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// Test static class to invoke from inside PowerShell script.
    /// </summary>
    public static class TestStaticClass
    {
        /// <summary>
        /// Sample static method to call from insider PowerShell script.
        /// </summary>
        /// <returns>String message</returns>
        public static string TestStaticMethod()
        {
            return "Hello, you have called the test static method.";
        }
    }

    /// <summary>
    /// Provides PowerShell script execution examples
    /// </summary>
    class PowerShellExecutor
    {
        /// <summary>
        /// Sample execution scenario 3: Namespace test
        /// </summary>
        /// <remarks>
        /// Executes a PowerShell script synchronously and utilizes classes in the callers namespace.
        /// </remarks>
        public void ExecuteSynchronouslyNamespaceTest()
        {
            using (PowerShell PowerShellInstance = PowerShell.Create())
            {
                // add a script that creates a new instance of an object from the caller's namespace
                PowerShellInstance.AddScript("$t = new-object PowerShellExecutionSample.TestObject;" +
                                             "$t.Name = 'created from inside PowerShell script'; $t;" +
                                             "$message = [PowerShellExecutionSample.TestStaticClass]::TestStaticMethod(); $message");

                // invoke execution on the pipeline (collecting output)
                Collection<PSObject> PSOutput = PowerShellInstance.Invoke();

                // loop through each output object item
                foreach (PSObject outputItem in PSOutput)
                {
                    if (outputItem != null)
                    {
                        Console.WriteLine(outputItem.BaseObject.GetType().FullName);

                        if (outputItem.BaseObject is TestObject)
                        {
                            TestObject testObj = outputItem.BaseObject as TestObject;
                            Console.WriteLine(testObj.Name);
                        }
                        else
                        {
                            Console.WriteLine(outputItem.BaseObject.ToString());
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Sample execution scenario 1: Synchronous
        /// </summary>
        /// <remarks>
        /// Executes a PowerShell script synchronously with input parameters and script output handling.
        /// </remarks>
        public void ExecuteSynchronously()
        {
            using (PowerShell PowerShellInstance = PowerShell.Create())
            {
                // use "AddScript" to add the contents of a script file to the end of the execution pipeline.
                // use "AddCommand" to add individual commands/cmdlets to the end of the execution pipeline.
                PowerShellInstance.AddScript("param($param1) $d = get-date; $s = 'test string value'; " +
                "$d; $s; $param1; get-service");

                // use "AddParameter" to add a single parameter to the last command/script on the pipeline.
                PowerShellInstance.AddParameter("param1", "parameter 1 value!");

                // invoke execution on the pipeline (collecting output)
                Collection<PSObject> PSOutput = PowerShellInstance.Invoke();

                // check the other output streams (for example, the error stream)
                if (PowerShellInstance.Streams.Error.Count > 0)
                {
                    // error records were written to the error stream.
                    // do something with the items found.
                }

                // loop through each output object item
                foreach (PSObject outputItem in PSOutput)
                {
                    // if null object was dumped to the pipeline during the script then a null
                    // object may be present here. check for null to prevent potential NRE.
                    if (outputItem != null)
                    {
                        //TODO: do something with the output item 
                        Console.WriteLine(outputItem.BaseObject.GetType().FullName);
                        Console.WriteLine(outputItem.BaseObject.ToString() + "\n");
                    }
                }
            }
        }

        /// <summary>
        /// Sample execution scenario 2: Asynchronous
        /// </summary>
        /// <remarks>
        /// Executes a PowerShell script asynchronously with script output and event handling.
        /// </remarks>
        public void ExecuteAsynchronously()
        {
            using (PowerShell PowerShellInstance = PowerShell.Create())
            {
                // this script has a sleep in it to simulate a long running script
                PowerShellInstance.AddScript("$s1 = 'test1'; $s2 = 'test2'; $s1; write-error 'some error';start-sleep -s 7; $s2");

                // prepare a new collection to store output stream objects
                PSDataCollection<PSObject> outputCollection = new PSDataCollection<PSObject>();
                outputCollection.DataAdded += outputCollection_DataAdded;

                // the streams (Error, Debug, Progress, etc) are available on the PowerShell instance.
                // we can review them during or after execution.
                // we can also be notified when a new item is written to the stream (like this):
                PowerShellInstance.Streams.Error.DataAdded += Error_DataAdded;

                // begin invoke execution on the pipeline
                // use this overload to specify an output stream buffer
                IAsyncResult result = PowerShellInstance.BeginInvoke<PSObject, PSObject>(null, outputCollection);

                // do something else until execution has completed.
                // this could be sleep/wait, or perhaps some other work
                while (result.IsCompleted == false)
                {
                    Console.WriteLine("Waiting for pipeline to finish...");
                    Thread.Sleep(1000);

                    // might want to place a timeout here...
                }

                Console.WriteLine("Execution has stopped. The pipeline state: " + PowerShellInstance.InvocationStateInfo.State);

                foreach (PSObject outputItem in outputCollection)
                {
                    //TODO: handle/process the output items if required
                    Console.WriteLine(outputItem.BaseObject.ToString());
                }
            }
        }

        /// <summary>
        /// Event handler for when data is added to the output stream.
        /// </summary>
        /// <param name="sender">Contains the complete PSDataCollection of all output items.</param>
        /// <param name="e">Contains the index ID of the added collection item and the ID of the PowerShell instance this event belongs to.</param>
        void outputCollection_DataAdded(object sender, DataAddedEventArgs e)
        {
            // do something when an object is written to the output stream
            Console.WriteLine("Object added to output.");
        }

        /// <summary>
        /// Event handler for when Data is added to the Error stream.
        /// </summary>
        /// <param name="sender">Contains the complete PSDataCollection of all error output items.</param>
        /// <param name="e">Contains the index ID of the added collection item and the ID of the PowerShell instance this event belongs to.</param>
        void Error_DataAdded(object sender, DataAddedEventArgs e)
        {
            // do something when an error is written to the error stream
            Console.WriteLine("An error was written to the Error stream!");
        }
    }
}
