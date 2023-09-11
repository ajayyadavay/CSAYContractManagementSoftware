using iText.Kernel.Pdf;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSAY_ContractManagementSoftware
{
    internal class CSAYNumToWord
    {
        public string ConvertNumberToNepaliWord(double number)
        {
            string nepaliWord="";

            double intNum = Math.Floor(number);
            double decimalPlace = Math.Round((number - intNum) * 100);
            string sDecimalPlace = NepaliWordOfNumber(decimalPlace);

            double numDivided = intNum / 100;
            intNum = Math.Floor(numDivided);
            double tens = Math.Round((numDivided - intNum) * 100);
            string sTens = NepaliWordOfNumber(tens);

            numDivided = intNum / 10;
            intNum = Math.Floor(numDivided);
            double hundred = Math.Round((numDivided - intNum) * 10);
            string sHundred = NepaliWordOfNumber(hundred);

            numDivided = intNum / 100;
            intNum = Math.Floor(numDivided);
            double thousand = Math.Round((numDivided - intNum) * 100);
            string sThousand = NepaliWordOfNumber(thousand);

            numDivided = intNum / 100;
            intNum = Math.Floor(numDivided);
            double lakh = Math.Round((numDivided - intNum) * 100);
            string sLakh = NepaliWordOfNumber(lakh);

            numDivided = intNum / 100;
            intNum = Math.Floor(numDivided);
            double crore = Math.Round((numDivided - intNum) * 100);
            string sCrore = NepaliWordOfNumber(crore);

            numDivided = intNum / 100;
            intNum = Math.Floor(numDivided);
            double arab = Math.Round((numDivided - intNum) * 100);
            string sArab = NepaliWordOfNumber(arab);

            numDivided = intNum / 100;
            intNum = Math.Floor(numDivided);
            double kharab = Math.Round((numDivided - intNum) * 100);
            string sKharab = NepaliWordOfNumber(kharab);

            string sKharab1 = string.IsNullOrEmpty(sKharab) ? "" : " " + sKharab + " " + "खरब"; // Kharab

            string sArab1 = string.IsNullOrEmpty(sArab) ? "" : " " + sArab + " " + "अरब"; // Arab
            string sCrore1 = string.IsNullOrEmpty(sCrore) ? "" : " " + sCrore + " " + "करोड"; // Crore
            string sLakh1 = string.IsNullOrEmpty(sLakh) ? "" : " " + sLakh + " " + "लाख"; // Lakh
            string sThousand1 = string.IsNullOrEmpty(sThousand) ? "" : " " + sThousand + " " + "हजार"; // Thousand
            string sHundred1 = string.IsNullOrEmpty(sHundred) ? "" : " " + sHundred + " " + "सय"; // Hundred
            string sTens1 = string.IsNullOrEmpty(sTens) ? "" : " " + sTens + " " + "रुपैयाँ"; // Rupees
            string sDecimalPlace1 = string.IsNullOrEmpty(sDecimalPlace) ? "" : " " + sDecimalPlace + " " + "पैसा"; // Paisa

            nepaliWord = sKharab1 + sArab1 + sCrore1 + sLakh1 + sThousand1 + sHundred1 + sTens1 + sDecimalPlace1 + " " + "मात्र";

            return nepaliWord;
        }

        private string NepaliWordOfNumber(double Number)
        {
            string nepali_num_word="";
            switch(Number)
            {
                case 1:
                    nepali_num_word = "एक";
                    break;
                case 2:
                    nepali_num_word = "दुई";
                        break;
                case 3:
                    nepali_num_word = "तीन";
                        break;
                case 4:
                nepali_num_word = "चार";
                    break;
                case 5:
                nepali_num_word = "पाँच";
                    break;
                case 6:
                nepali_num_word = "छ";
                    break;
                case 7:
                nepali_num_word = "सात";
                    break;
                case 8:
                nepali_num_word = "आठ";
                    break;
                case 9:
                nepali_num_word = "नौ";
                    break;
                case 10:
                nepali_num_word = "दश";
                    break;

                case 11:
                nepali_num_word = "एघार";
                    break;
                case 12:
                nepali_num_word = "बाह्र";
                    break;
                case 13:
                    nepali_num_word = "तेह्र";
                    break;
                case 14:
                    nepali_num_word = "चौध";
                    break;
                case 15:
                    nepali_num_word = "पन्ध्र";
                    break;
                case 16:
                    nepali_num_word = "सोह्र";
                    break;
                case 17:
                    nepali_num_word = "सत्र";
                    break;
                case 18:
                    nepali_num_word = "अठार";
                    break;
                case 19:
                    nepali_num_word = "उन्नाइस";
                    break;
                case 20:
                    nepali_num_word = "बिस";
                    break;

                case 21:
                    nepali_num_word = "एक्काइस";
                    break;
                case 22:
                    nepali_num_word = "बाइस";
                    break;
                case 23:
                    nepali_num_word = "तेइस";
                    break;
                case 24:
                    nepali_num_word = "चौबिस";
                    break;
                case 25:
                    nepali_num_word = "पच्चिस";
                    break;
                case 26:
                    nepali_num_word = "छब्बिस";
                    break;
                case 27:
                    nepali_num_word = "सत्ताइस";
                    break;
                case 28:
                    nepali_num_word = "अट्ठाइस";
                    break;
                case 29:
                    nepali_num_word = "उनन्तिस";
                    break;
                case 30:
                    nepali_num_word = "तिस";
                    break;

                case 31:
                    nepali_num_word = "एकतिस";
                    break;
                case 32:
                    nepali_num_word = "बत्तिस";
                    break;
                case 33:
                    nepali_num_word = "तेत्तिस";
                    break;
                case 34:
                    nepali_num_word = "चौतिस";
                    break;
                case 35:
                    nepali_num_word = "पैँतिस";
                    break;
                case 36:
                    nepali_num_word = "छत्तिस";
                    break;
                case 37:
                    nepali_num_word = "सैँतिस";
                    break;
                case 38:
                    nepali_num_word = "अठतिस";
                    break;
                case 39:
                    nepali_num_word = "उनन्चालिस";
                    break;
                case 40:
                    nepali_num_word = "चालिस";
                    break;

                case 41:
                    nepali_num_word = "एकचालिस";
                    break;
                case 42:
                    nepali_num_word = "बयालिस";
                    break;
                case 43:
                    nepali_num_word = "त्रिचालिस";
                    break;
                case 44:
                    nepali_num_word = "चवालिस";
                    break;
                case 45:
                    nepali_num_word = "पैँतालिस";
                    break;
                case 46:
                    nepali_num_word = "छयालिस";
                    break;
                case 47:
                    nepali_num_word = "सतचालिस";
                    break;
                case 48:
                    nepali_num_word = "अठचालिस";
                    break;
                case 49:
                    nepali_num_word = "उनन्चास";
                    break;
                case 50:
                    nepali_num_word = "पचास";
                    break;

                case 51:
                    nepali_num_word = "एकाउन्न";
                    break;
                case 52:
                    nepali_num_word = "बाउन्न";
                    break;
                case 53:
                    nepali_num_word = "त्रिपन्न";
                    break;
                case 54:
                    nepali_num_word = "चवन्न";
                    break;
                case 55:
                    nepali_num_word = "पचपन्न";
                    break;
                case 56:
                    nepali_num_word = "छपन्न";
                    break;
                case 57:
                    nepali_num_word = "सन्ताउन्न";
                    break;
                case 58:
                    nepali_num_word = "अन्ठाउन्न";
                    break;
                case 59:
                    nepali_num_word = "उनसट्ठी";
                    break;
                case 60:
                    nepali_num_word = "साठी";
                    break;

                case 61:
                    nepali_num_word = "एकसट्ठी";
                    break;
                case 62:
                    nepali_num_word = "बयसट्ठी";
                    break;
                case 63:
                    nepali_num_word = "त्रिसट्ठी";
                    break;
                case 64:
                    nepali_num_word = "चौसट्ठी";
                    break;
                case 65:
                    nepali_num_word = "पैँसट्ठी";
                    break;
                case 66:
                    nepali_num_word = "छयसट्ठी";
                    break;
                case 67:
                    nepali_num_word = "सतसट्ठी";
                    break;
                case 68:
                    nepali_num_word = "अठसट्ठी";
                    break;
                case 69:
                    nepali_num_word = "उनन्सत्तरी";
                    break;
                case 70:
                    nepali_num_word = "सत्तरी";
                    break;

                case 71:
                    nepali_num_word = "एकहत्तर";
                    break;
                case 72:
                    nepali_num_word = "बहत्तर";
                    break;
                case 73:
                    nepali_num_word = "त्रिहत्तर";
                    break;
                case 74:
                    nepali_num_word = "चौहत्तर";
                    break;
                case 75:
                    nepali_num_word = "पचहत्तर";
                    break;
                case 76:
                    nepali_num_word = "छयहत्तर";
                    break;
                case 77:
                    nepali_num_word = "सतहत्तर";
                    break;
                case 78:
                    nepali_num_word = "अठहत्तर";
                    break;
                case 79:
                    nepali_num_word = "उनासी";
                    break;
                case 80:
                    nepali_num_word = "असी";
                    break;

                case 81:
                    nepali_num_word = "एकासी";
                    break;
                case 82:
                    nepali_num_word = "बयासी";
                    break;
                case 83:
                    nepali_num_word = "त्रियासी";
                    break;
                case 84:
                    nepali_num_word = "चौरासी";
                    break;
                case 85:
                    nepali_num_word = "पचासी";
                    break;
                case 86:
                    nepali_num_word = "छयासी";
                    break;
                case 87:
                    nepali_num_word = "सतासी";
                    break;
                case 88:
                    nepali_num_word = "अठासी";
                    break;
                case 89:
                    nepali_num_word = "उनान्नब्बे";
                    break;
                case 90:
                    nepali_num_word = "नब्बे";
                    break;

                case 91:
                    nepali_num_word = "एकान्नब्बे";
                    break;
                case 92:
                    nepali_num_word = "बयान्नब्बे";
                    break;
                case 93:
                    nepali_num_word = "त्रियान्नब्बे";
                    break;
                case 94:
                    nepali_num_word = "चौरान्नब्बे";
                    break;
                case 95:
                    nepali_num_word = "पन्चान्नब्बे";
                    break;
                case 96:
                    nepali_num_word = "छयान्नब्बे";
                    break;
                case 97:
                    nepali_num_word = "सन्तान्नब्बे";
                    break;
                case 98:
                    nepali_num_word = "अन्ठान्नब्बे";
                    break;
                case 99:
                    nepali_num_word = "उनान्सय";
                    break;

                default:
                    nepali_num_word = "";
                    break;
            }
            return nepali_num_word;
        }


        public string ConvertNumberToEnglishWord(double number)
        {
            string EnglishWord = "";

            double intNum = Math.Floor(number);
            double decimalPlace = Math.Round((number - intNum) * 100);
            string sDecimalPlace = EnglishWordOfNumber(decimalPlace);

            double numDivided = intNum / 100;
            intNum = Math.Floor(numDivided);
            double tens = Math.Round((numDivided - intNum) * 100);
            string sTens = EnglishWordOfNumber(tens);

            numDivided = intNum / 10;
            intNum = Math.Floor(numDivided);
            double hundred = Math.Round((numDivided - intNum) * 10);
            string sHundred = EnglishWordOfNumber(hundred);

            numDivided = intNum / 100;
            intNum = Math.Floor(numDivided);
            double thousand = Math.Round((numDivided - intNum) * 100);
            string sThousand = EnglishWordOfNumber(thousand);

            numDivided = intNum / 100;
            intNum = Math.Floor(numDivided);
            double lakh = Math.Round((numDivided - intNum) * 100);
            string sLakh = EnglishWordOfNumber(lakh);

            numDivided = intNum / 100;
            intNum = Math.Floor(numDivided);
            double crore = Math.Round((numDivided - intNum) * 100);
            string sCrore = EnglishWordOfNumber(crore);

            numDivided = intNum / 100;
            intNum = Math.Floor(numDivided);
            double arab = Math.Round((numDivided - intNum) * 100);
            string sArab = EnglishWordOfNumber(arab);

            numDivided = intNum / 100;
            intNum = Math.Floor(numDivided);
            double kharab = Math.Round((numDivided - intNum) * 100);
            string sKharab = EnglishWordOfNumber(kharab);

            string sKharab1 = string.IsNullOrEmpty(sKharab) ? "" : " " + sKharab + " " + "Kharab"; // Kharab

            string sArab1 = string.IsNullOrEmpty(sArab) ? "" : " " + sArab + " " + "Arab"; // Arab
            string sCrore1 = string.IsNullOrEmpty(sCrore) ? "" : " " + sCrore + " " + "Crore"; // Crore
            string sLakh1 = string.IsNullOrEmpty(sLakh) ? "" : " " + sLakh + " " + "Lakh"; // Lakh
            string sThousand1 = string.IsNullOrEmpty(sThousand) ? "" : " " + sThousand + " " + "Thousand"; // Thousand
            string sHundred1 = string.IsNullOrEmpty(sHundred) ? "" : " " + sHundred + " " + "Hundred"; // Hundred
            string sTens1 = string.IsNullOrEmpty(sTens) ? "" : " " + sTens + " " + "Rupees"; // Rupees
            string sDecimalPlace1 = string.IsNullOrEmpty(sDecimalPlace) ? "" : " " + sDecimalPlace + " " + "Paisa"; // Paisa

            EnglishWord = sKharab1 + sArab1 + sCrore1 + sLakh1 + sThousand1 + sHundred1 + sTens1 + sDecimalPlace1 + " " + "Only";

            return EnglishWord;
        }

        private string EnglishWordOfNumber(double Number)
        {
            string English_num_word = "";
            switch (Number)
            {
                case 1:
                    English_num_word = "One";
                    break;
                case 2:
                    English_num_word = "Two";
                    break;
                case 3:
                    English_num_word = "Three";
                    break;
                case 4:
                    English_num_word = "Four";
                    break;
                case 5:
                    English_num_word = "Five";
                    break;
                case 6:
                    English_num_word = "Six";
                    break;
                case 7:
                    English_num_word = "Seven";
                    break;
                case 8:
                    English_num_word = "Eight";
                    break;
                case 9:
                    English_num_word = "Nine";
                    break;
                case 10:
                    English_num_word = "Ten";
                    break;

                case 11:
                    English_num_word = "Eleven";
                    break;
                case 12:
                    English_num_word = "Twelve";
                    break;
                case 13:
                    English_num_word = "Thirteen";
                    break;
                case 14:
                    English_num_word = "Fourteen";
                    break;
                case 15:
                    English_num_word = "Fifteen";
                    break;
                case 16:
                    English_num_word = "Sixteen";
                    break;
                case 17:
                    English_num_word = "Seventeen";
                    break;
                case 18:
                    English_num_word = "Eighteen";
                    break;
                case 19:
                    English_num_word = "Nineteen";
                    break;
                case 20:
                    English_num_word = "Twenty";
                    break;

                case 21:
                    English_num_word = "Twenty One";
                    break;
                case 22:
                    English_num_word = "Twenty Two";
                    break;
                case 23:
                    English_num_word = "Twenty Three";
                    break;
                case 24:
                    English_num_word = "Twenty Four";
                    break;
                case 25:
                    English_num_word = "Twenty Five";
                    break;
                case 26:
                    English_num_word = "Twenty Six";
                    break;
                case 27:
                    English_num_word = "Twenty Seven";
                    break;
                case 28:
                    English_num_word = "Twenty Eight";
                    break;
                case 29:
                    English_num_word = "Twenty Nine";
                    break;
                case 30:
                    English_num_word = "Thirty";
                    break;

                case 31:
                    English_num_word = "Thirty One";
                    break;
                case 32:
                    English_num_word = "Thirty Two";
                    break;
                case 33:
                    English_num_word = "Thirty Three";
                    break;
                case 34:
                    English_num_word = "Thirty Four";
                    break;
                case 35:
                    English_num_word = "Thirty Five";
                    break;
                case 36:
                    English_num_word = "Thirty Six";
                    break;
                case 37:
                    English_num_word = "Thirty Seven";
                    break;
                case 38:
                    English_num_word = "Thirty Eight";
                    break;
                case 39:
                    English_num_word = "Thirty Nine";
                    break;
                case 40:
                    English_num_word = "Forty";
                    break;

                case 41:
                    English_num_word = "Forty One";
                    break;
                case 42:
                    English_num_word = "Forty Two";
                    break;
                case 43:
                    English_num_word = "Forty Three";
                    break;
                case 44:
                    English_num_word = "Forty Four";
                    break;
                case 45:
                    English_num_word = "Forty Five";
                    break;
                case 46:
                    English_num_word = "Forty Six";
                    break;
                case 47:
                    English_num_word = "Forty Seven";
                    break;
                case 48:
                    English_num_word = "Forty Eight";
                    break;
                case 49:
                    English_num_word = "Forty Nine";
                    break;
                case 50:
                    English_num_word = "Fifty";
                    break;

                case 51:
                    English_num_word = "Fifty One";
                    break;
                case 52:
                    English_num_word = "Fifty Two";
                    break;
                case 53:
                    English_num_word = "Fifty Three";
                    break;
                case 54:
                    English_num_word = "Fifty Four";
                    break;
                case 55:
                    English_num_word = "Fifty Five";
                    break;
                case 56:
                    English_num_word = "Fifty Six";
                    break;
                case 57:
                    English_num_word = "Fifty Seven";
                    break;
                case 58:
                    English_num_word = "Fifty Eight";
                    break;
                case 59:
                    English_num_word = "Fifty Nine";
                    break;
                case 60:
                    English_num_word = "Sixty";
                    break;

                case 61:
                    English_num_word = "Sixty One";
                    break;
                case 62:
                    English_num_word = "Sixty Two";
                    break;
                case 63:
                    English_num_word = "Sixty Three";
                    break;
                case 64:
                    English_num_word = "Sixty Four";
                    break;
                case 65:
                    English_num_word = "Sixty Five";
                    break;
                case 66:
                    English_num_word = "Sixty Six";
                    break;
                case 67:
                    English_num_word = "Sixty Seven";
                    break;
                case 68:
                    English_num_word = "Sixty Eight";
                    break;
                case 69:
                    English_num_word = "Sixty Nine";
                    break;
                case 70:
                    English_num_word = "Seventy";
                    break;

                case 71:
                    English_num_word = "Seventy One";
                    break;
                case 72:
                    English_num_word = "Seventy Two";
                    break;
                case 73:
                    English_num_word = "Seventy Three";
                    break;
                case 74:
                    English_num_word = "Seventy Four";
                    break;
                case 75:
                    English_num_word = "Seventy Five";
                    break;
                case 76:
                    English_num_word = "Seventy Six";
                    break;
                case 77:
                    English_num_word = "Seventy Seven";
                    break;
                case 78:
                    English_num_word = "Seventy Eight";
                    break;
                case 79:
                    English_num_word = "Seventy Nine";
                    break;
                case 80:
                    English_num_word = "Eighty";
                    break;

                case 81:
                    English_num_word = "Eighty One";
                    break;
                case 82:
                    English_num_word = "Eighty Two";
                    break;
                case 83:
                    English_num_word = "Eighty Three";
                    break;
                case 84:
                    English_num_word = "Eighty Four";
                    break;
                case 85:
                    English_num_word = "Eighty Five";
                    break;
                case 86:
                    English_num_word = "Eighty Six";
                    break;
                case 87:
                    English_num_word = "Eighty Seven";
                    break;
                case 88:
                    English_num_word = "Eighty Eight";
                    break;
                case 89:
                    English_num_word = "Eighty Nine";
                    break;
                case 90:
                    English_num_word = "Ninety";
                    break;

                case 91:
                    English_num_word = "Ninety One";
                    break;
                case 92:
                    English_num_word = "Ninety Two";
                    break;
                case 93:
                    English_num_word = "Ninety Three";
                    break;
                case 94:
                    English_num_word = "Ninety Four";
                    break;
                case 95:
                    English_num_word = "Ninety Five";
                    break;
                case 96:
                    English_num_word = "Ninety Six";
                    break;
                case 97:
                    English_num_word = "Ninety Seven";
                    break;
                case 98:
                    English_num_word = "Ninety Eight";
                    break;
                case 99:
                    English_num_word = "Ninety Nine";
                    break;

                default:
                    English_num_word = "";
                    break;
            }
            return English_num_word;
        }

    }
}
