using System;
using System.Collections.Generic;

namespace ASConverter {
    public class AccountSectionSorterByStartDate : IComparer<AccountSection> {
        public int Compare(AccountSection x, AccountSection y) {
            if (x == y) {
                return 0;
            }

            return x.StartData < y.StartData ? -1 : 1;
        }
    }

    public class AccountSectionSorterByEndDate : IComparer<AccountSection> {
        public int Compare(AccountSection x, AccountSection y) {
            if (x == y) {
                return 0;
            }

            return x.EndData < y.EndData? -1 : 1;
        }
    }
}