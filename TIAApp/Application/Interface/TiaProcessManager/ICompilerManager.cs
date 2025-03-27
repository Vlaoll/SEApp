// Ignore Spelling: Conf

using Siemens.Engineering;

namespace seConfSW.Services
{
    /// <summary>
    /// Interface for the CompilerManager service that handles compilation of devices in a TIA project.
    /// </summary>
    public interface ICompilerManager
    {
        /// <summary>
        /// Compiles the specified device in the given TIA project.
        /// </summary>
        /// <param name="devName">Name of the device to compile.</param>
        /// <param name="project">The TIA project containing the device.</param>
        /// <exception cref="ArgumentNullException">Thrown when devName is null or empty.</exception>
        void Compile(string devName, Project project);
    }
}