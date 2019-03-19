
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Contensive.BaseClasses;


namespace Contensive.addonManager
{
    public class addonManagerSelectorClass : AddonBaseClass
    {
        //
        // class scope
        //
        private const string guidAddonManagerActiveX = "{1DC06F61-1837-419B-AF36-D5CC41E1C9FD}";
        //
        //
        //
        public override object Execute(CPBaseClass cp)
        {
            if (string.Compare( cp.Version, "5") < 0)
            {
                cp.Utils.ExecuteAddon(guidAddonManagerActiveX);
            }
            else
            {
                Contensive.addonManager.addonManager5Class addonManager = new addonManager5Class();
                return addonManager.Execute(cp);

            }
            return string.Empty;
        }

    }
}
