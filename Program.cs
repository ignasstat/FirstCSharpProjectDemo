using System;
using PointLibrary;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace Prepare4Test
{

// … skaldome pavyzdžius funkcijomis
public class SamplesArray
        {
            public static void Main()
            {
                //IllustrateArrayAsReadOnly();
                //IllustrateClone();
                //IllustrateForeach();
                IllustrateBinarySearch();
            }

            static void IllustrateArrayAsReadOnly()
            {
                string[] myArr = { "The", "quick", "brown", "fox" };
                Console.WriteLine("The string array initially contains the following values:");
                IList<string> myList = Array.AsReadOnly(myArr);
                Console.WriteLine("The read-only IList contains the following values:");

                // Attempt to change a value through the wrapper.
                myList[3] = "CAT";

                // Change a value in the original array.
                myArr[2] = "RED";
            }

            static void IllustrateClone()
            {
                string[] myArr1 = { "The", "quick", "brown", "fox" };
                string[] myArr2 = (string[])(myArr1.Clone());
                //string[] myArr2 = myArr1;

                Console.WriteLine(string.Join(", ", myArr1));
                myArr1[3] = "wolf";
                Console.WriteLine(string.Join(", ", myArr1));

                Console.WriteLine(string.Join(", ", myArr2));
            }

            static void IllustrateForeach()
            {
                string[] myArr1 = { "The", "quick", "brown", "fox" };
                Console.WriteLine(string.Join(", ", myArr1));
                Array.ForEach(myArr1, word => {
                    Console.Write(word.ToUpper() + ", ");
                });
                Console.WriteLine("\n" + string.Join(", ", myArr1));
            }

            static void IllustrateBinarySearch()
            {
                string[] myArr1 = { "The", "quick", "brown", "fox" };
                Console.WriteLine("Before sorting: " + string.Join(", ", myArr1));
                Array.Sort(myArr1);
                Console.WriteLine("After sorting: " + string.Join(", ", myArr1));

                int idx = Array.BinarySearch(myArr1, "quick");
  //      Checking if idx more than 0, if yes ten print myArr1[idx, if not the "Sor.."]
                Console.WriteLine(idx > 0 ? myArr1[idx] : "Sorry, not found!");
            }
        }                                                                                         

  }





