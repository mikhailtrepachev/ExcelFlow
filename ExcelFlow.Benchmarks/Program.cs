using BenchmarkDotNet.Running;

BenchmarkSwitcher.FromAssembly(typeof(ExcelFlow.Benchmarks.BenchRow).Assembly).Run(args);
