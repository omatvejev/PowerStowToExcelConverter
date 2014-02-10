using System;
using System.Collections.Generic;

namespace PowerStowToExcelConverter.Model
{
    class Terminal
    {
        private string name;
        private string vessel;
        private string voyage;

        public Terminal(string name, string vessel, string voyage)
        {
            this.name = name;
            this.vessel = vessel;
            this.voyage = voyage;
        }

        public string Name
        {
            get
            {
                return this.name;
            }
        }

        public string Voyage
        {
            get
            {
                return this.voyage;
            }
        }

        public string Vessel
        {
            get
            {
                return this.vessel;
            }
        }
    }
}
