using System;

namespace OneAssist.Assistant.Services
{
    public class CorporateGovernancePolicy
    {
        public bool IsBrandLocked { get; set; } = false;
        public bool IsDictionaryLocked { get; set; } = false;
        public bool IsRulesLocked { get; set; } = false;
        public string OrganizationName { get; set; } = "Enterprise Standard";
        public string PolicyEnforcedDate { get; set; } = DateTime.Now.ToString("yyyy-MM-dd");
    }

    public static class AdminGovernanceService
    {
        private static CorporateGovernancePolicy _currentPolicy = new CorporateGovernancePolicy();

        public static CorporateGovernancePolicy CurrentPolicy => _currentPolicy;

        public static void SetPolicy(bool lockBrand, bool lockDict, bool lockRules, string orgName)
        {
            _currentPolicy.IsBrandLocked = lockBrand;
            _currentPolicy.IsDictionaryLocked = lockDict;
            _currentPolicy.IsRulesLocked = lockRules;
            _currentPolicy.OrganizationName = orgName;
        }

        public static bool CanUserModifyBrand() => !_currentPolicy.IsBrandLocked;
        public static bool CanUserModifyDictionary() => !_currentPolicy.IsDictionaryLocked;
        public static bool CanUserModifyRules() => !_currentPolicy.IsRulesLocked;
    }
}
