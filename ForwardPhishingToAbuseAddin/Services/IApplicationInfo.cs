namespace ForwardPhishingToAbuseAddin.Services
{
	public interface IApplicationInfo
	{
		string ApplicationCompany { get; }
		string ApplicationProduct { get; }

		string ApplicationVersion { get; }
	}
}