# Update3ePresentations
**Demonstrates how to update 3e objects without the IDE or U3E**

During an upgrade of TRE Elite 3e from 2.5 to 2.7, it became clear that the stock conversion of ReportLayouts to Presentations didn't work quite as well as hoped.

This utility was used to reset the metric run ID back to its original value and to set the Distinct Layouts attribute on the converted presentations.

As such, it demonstrates a couple of potentially useful techniques:
- How to run OQL queries in your own .net application (see GetCustomPresentationsProcess.vb)
- How to safely perform a mass update of 3e object definitions using the 3e framework (see UpdatePresentationsProcess.vb)

Both these tasks are achieved by creating a separate AppDomain where 3e can run happily using the framework DLLs from the Utilities3e directory and the application DLLs from the Staging directory.
