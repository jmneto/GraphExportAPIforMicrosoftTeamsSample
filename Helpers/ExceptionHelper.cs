//MIT License
//
//Copyright (c) 2024 Microsoft - Jose Batista-Neto.
//
//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:
//
//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.
//
//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.

namespace GraphExportAPIforMicrosoftTeamsSample.Helpers;

// Simple exception class to handle fatal exceptions of the Type IO
// I used this to create extended exceptions when we cannot access the storage account
// in this case we just print the error messages to the console and quit the application
internal class IOFatalException : Exception
{
    public IOFatalException(string message) : base(message)
    {
    }

    public IOFatalException(string message, Exception innerException) : base(message, innerException)
    {
    }
}

internal class OptimisticConcurrencyException : Exception
{
    public OptimisticConcurrencyException(string message) : base(message)
    {
    }

    public OptimisticConcurrencyException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
