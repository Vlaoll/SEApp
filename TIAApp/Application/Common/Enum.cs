// Ignore Spelling: Conf

namespace seConfSW.Services
{
    /// <summary>
    /// Represents the status of the connection to TIA Portal.
    /// </summary>
    public enum ConnectionStatus
    {
        /// <summary>
        /// No running instance of TIA Portal was found.
        /// </summary>
        NoInstanceFound,

        /// <summary>
        /// Connected to TIA Portal, but no project was found.
        /// </summary>
        NoProjectFound,

        /// <summary>
        /// Successfully connected to TIA Portal and a project.
        /// </summary>
        ConnectedSuccessfully,

        /// <summary>
        /// More than one running instance of TIA Portal was found.
        /// </summary>
        MultipleInstancesFound
    }
}