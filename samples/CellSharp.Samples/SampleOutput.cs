namespace CellSharp.Samples;

internal static class SampleOutput
{
    internal static string Prepare()
    {
        var repositoryRoot = FindRepositoryRoot();
        var outputDirectory = Path.Combine(repositoryRoot, "samples", "CellSharp.Samples", "output");

        Directory.CreateDirectory(outputDirectory);
        // This is a fixed, sample-owned directory, never an inferred user path.
        // Keep .gitkeep so running the samples does not dirty a checkout.
        foreach (var file in Directory.GetFiles(outputDirectory))
        {
            if (!string.Equals(Path.GetFileName(file), ".gitkeep", StringComparison.Ordinal))
            {
                File.Delete(file);
            }
        }

        foreach (var directory in Directory.GetDirectories(outputDirectory))
        {
            Directory.Delete(directory, recursive: true);
        }

        return outputDirectory;
    }

    private static string FindRepositoryRoot()
    {
        for (DirectoryInfo? directory = new DirectoryInfo(Directory.GetCurrentDirectory()); directory is not null; directory = directory.Parent)
        {
            if (File.Exists(Path.Combine(directory.FullName, "CellSharp.sln")))
            {
                return directory.FullName;
            }
        }

        throw new InvalidOperationException("Run the samples from a checkout of the CellSharp repository.");
    }
}
