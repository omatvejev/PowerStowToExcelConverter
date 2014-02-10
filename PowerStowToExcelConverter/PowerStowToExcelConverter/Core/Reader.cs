using System;
using System.IO;
using System.Collections;

namespace PowerStowToExcelConverter.Core
{
    class Reader
    {
        private String filename;
        public Reader(String filename) 
        {
            this.filename = filename;
        }

        // Parses the file and returns the name of the terminal, 
        public string parseTerminalInformation()
        {
            string output = null;
            try
            {
                using (TextReader tr = new StreamReader(this.filename))
                {
                    string line = null;
                    while ((line = tr.ReadLine()) != null)
                    {
                        if (line.Contains("PORT AND TERMINAL:"))
                        {
                            output = line.Substring(19);

                            // remove unnecessary text
                            output = output.TrimEnd(',');
                            return output;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return output;
        }

        // Parses the file and returns the name of the vessel and voyage number
        public string[] parseVesselandVoyageInformation()
        {
            string input = null;
            try
            {
                using (TextReader tr = new StreamReader(this.filename))
                {
                    string line = null;
                    while ((line = tr.ReadLine()) != null)
                    {
                        if (line.Contains("VESSEL AND VOYAGE:"))
                        {
                            input = line.Substring(19);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            string[] output = null;
            if (!(input == null))
            {
                // There are 10 spaces between the vessel name and voyage. Split accordingly
                output = input.Split(new string[] { "          " }, StringSplitOptions.None);
            }

            return output;
        }

        public string[] parseCompaniesInformation()
        {
            string[] companies = null;

            try
            {
                using (TextReader tr = new StreamReader(this.filename))
                {
                    string line = null;
                    while ((line = tr.ReadLine()) != null)
                    {
                        if (line.Contains("6.7 SLOT UTILIZATION SUMMARY"))
                        {
                            // Skip next line
                            tr.ReadLine();
                            ArrayList list = new ArrayList();

                            // Continue reading next lines until the file end or we reach an empty line
                            while ((line = tr.ReadLine()) != null && !line.Equals(""))
                            {
                                // Break the string by space
                                string[] temp = line.Split(' ');
                                list.Add(temp[0]);
                            }
                            // Convert the list to array
                            companies = list.ToArray(typeof(string)) as string[];
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return companies;
        }

        public string[] parseShipmentPorts()
        {
            // Stores all the ports
            ArrayList portName = new ArrayList();

            try
            {
                using (TextReader tr = new StreamReader(this.filename))
                {
                    string line = null;

                    while ((line = tr.ReadLine()) != null)
                    {
                        // Read all the ports in the full containers section
                        if (line.Contains("6.1 FULL CONTAINERS"))
                        {
                            // Go to next line
                            line = tr.ReadLine();

                            // Read the line until we reach the next empty line. Also, make sure that there is a line
                            // in case the file has some data problems
                            while ((line = tr.ReadLine()) != null && !line.Equals(""))
                            {
                                // Disregard the total line
                                if (line.Contains("Total"))
                                    continue;

                                string port = "";
                                int i = 0;

                                // Loop through the line until we find a character that is not a letter which indicates that
                                // the port name is complete
                                while (i < line.Length && (line[i] < 48 || line[i] > 57))
                                {
                                    port = port + line[i];
                                    i++;
                                }

                                portName.Add(port);
                            }
                        }

                        // Read all the ports in the full containers section
                        if (line.Contains("6.4 EMPTY CONTAINERS"))
                        {
                            // Go to next line
                            line = tr.ReadLine();

                            // Read the line until we reach the next empty line. Also, make sure that there is a line
                            // in case the file has some data problems
                            while ((line = tr.ReadLine()) != null && !line.Equals(""))
                            {
                                // Disregard the total line
                                if (line.Contains("Total"))
                                    continue;

                                string port = "";
                                int i = 0;

                                // Loop through the line until we find a character that is not a letter which indicates that
                                // the port name is complete
                                while (i < line.Length && (line[i] < 48 || line[i] > 57))
                                {
                                    port = port + line[i];
                                    i++;
                                }

                                // Make sure this port doesn't exist
                                if (!portName.Contains(port))
                                    portName.Add(port);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            // Returns the string of the portName might be empty!
            return portName.ToArray(typeof(string)) as string[];
        }

        public string[][] parseFullContainers(string company)
        {
            string[][] containerData = null;
            try
            {
                using (TextReader tr = new StreamReader(this.filename))
                {
                    string line = null;
                    while ((line = tr.ReadLine()) != null)
                    {
                        if (line.Contains("6.1 FULL CONTAINERS"))
                        {
                            // Go to next line
                            line = tr.ReadLine();

                            // Break the string by space 
                            string[] temp = line.Split(' ');

                            ArrayList companyList = new ArrayList();

                            // Determine the companies that have full containers
                            foreach (string str in temp)
                            {
                                if (!str.Equals(""))
                                    companyList.Add(str);
                            }

                            int position = -1;
                            for (int i = 0; i < companyList.Count; i++)
                            {
                                // we found the company we are currently interested in
                                if (companyList[i].Equals(company))
                                {
                                    position = i;
                                    break;
                                }
                            }

                            // The company was found
                            if (position != -1)
                            {
                                ArrayList portName = new ArrayList();
                                ArrayList containers = new ArrayList();
                                while ((line = tr.ReadLine()) != null && !line.Equals("") && !line.Contains("Total"))
                                {
                                    int i = 0;
                                    string s = "";

                                    // Continue to loop over the string until we reach the end of the line or we find a number
                                    // to generate the port that the container is being shipped to
                                    while (i < line.Length && (line[i] < 48 || line[i] > 57))
                                    {
                                        s = s + line[i];
                                        i++;
                                    }
                                    portName.Add(s);

                                    // No longer need the port name
                                    line = line.Substring(i);

                                    // Split the line into tokens
                                    string[] strArray = line.Split(' ');

                                    // Keeps track of the current column
                                    int j = 0;

                                    // Loop through each token and remove all the empty space
                                    foreach (string str in strArray)
                                    {
                                        // Do not include empty spaces and only add the containers in the right column
                                        if (!str.Equals(""))
                                        {
                                            // The column equals to the position of the company, therefore we have the companies container
                                            // information for the specific port
                                            if (j == position)
                                            {
                                                containers.Add(str);
                                            }
                                            j++;
                                        }
                                    }
                                }
                                // Convert the lists into the arrays
                                containerData = new string[2][];
                                containerData[0] = portName.ToArray(typeof(string)) as string[];
                                containerData[1] = containers.ToArray(typeof(string)) as string[];
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return containerData;
        }

        public string[][] parseEmptyContainers(string company)
        {
            string[][] containerData = null;
            try
            {
                using (TextReader tr = new StreamReader(this.filename))
                {
                    string line = null;
                    while ((line = tr.ReadLine()) != null)
                    {
                        if (line.Contains("6.4 EMPTY CONTAINERS"))
                        {
                            // Go to next line
                            line = tr.ReadLine();

                            // Break the string by space 
                            string[] temp = line.Split(' ');

                            ArrayList companyList = new ArrayList();

                            // Determine the companies that have full containers
                            foreach (string str in temp)
                            {
                                if (!str.Equals(""))
                                    companyList.Add(str);
                            }

                            int position = -1;
                            for (int i = 0; i < companyList.Count; i++)
                            {
                                // we found the company we are currently interested in
                                if (companyList[i].Equals(company))
                                {
                                    position = i;
                                    break;
                                }
                            }

                            // The company was found
                            if (position != -1)
                            {
                                ArrayList portName = new ArrayList();
                                ArrayList containers = new ArrayList();
                                while ((line = tr.ReadLine()) != null && !line.Equals("") && !line.Contains("Total"))
                                {
                                    int i = 0;
                                    string s = "";

                                    // Continue to loop over the string until we reach the end of the line or we find a number
                                    // to generate the port that the container is being shipped to
                                    while (i < line.Length && (line[i] < 48 || line[i] > 57))
                                    {
                                        s = s + line[i];
                                        i++;
                                    }
                                    portName.Add(s);

                                    // No longer need the port name
                                    line = line.Substring(i);

                                    // Split the line into tokens
                                    string[] strArray = line.Split(' ');

                                    // Keeps track of the current column
                                    int j = 0;

                                    // Loop through each token and remove all the empty space
                                    foreach (string str in strArray)
                                    {
                                        // Do not include empty spaces and only add the containers in the right column
                                        if (!str.Equals(""))
                                        {
                                            // The column equals to the position of the company, therefore we have the companies container
                                            // information for the specific port
                                            if (j == position)
                                            {
                                                containers.Add(str);
                                            }
                                            j++;
                                        }
                                    }
                                }
                                // Convert the lists into the arrays
                                containerData = new string[2][];
                                containerData[0] = portName.ToArray(typeof(string)) as string[];
                                containerData[1] = containers.ToArray(typeof(string)) as string[];
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return containerData;
        }
    }
}
