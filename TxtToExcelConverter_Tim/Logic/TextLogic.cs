using System;
using System.Collections.Generic;
using System.Linq;
using TxtToExcelConverter_Tim.Enumerations;
using TxtToExcelConverter_Tim.Models;

namespace TxtToExcelConverter_Tim.Logic
{
    public static class TextLogic
    {
        public static TableModel[] GetModelsFromString(string text)
        {
            string fromString = "-\n\n";

            // чтобы не использовать на других строках случайно
            if (!text.Contains(fromString))
                throw new Exception("Передан неподходящий тип строки");

            text = text.Substring(text.IndexOf(fromString) + fromString.Length);

            if (text.Contains("NetName"))
                text = text.Remove(text.IndexOf("NetName"));

            string[] strings = text.Split('\n');

            List<string> strList = new List<string>();

            #region Получение целых строк (удаление переносов)

            for (int i = 0; i < strings.Length; i++)
            {
                string str = strings[i];

                if (string.IsNullOrEmpty(str))
                    continue;

                #region ИЗБАВЛЕНИЕ ОТ ЕБАНЫХ ПУСТЫХ СТРОК

                bool isNotWhiteSpaces = false;
                
                foreach (char ch in str)
                {
                    if (!char.IsWhiteSpace(ch))
                    {
                        isNotWhiteSpaces = true;
                    }
                }
                
                if (isNotWhiteSpaces == false)
                    continue;

                #endregion

                if (str[0] == ' ')
                {
                    str = str.Replace(" ", string.Empty);

                    strList.Remove(strList.Last());

                    strList.Add(strings[i - 1] + str);
                }
                else
                {
                    strList.Add(str);
                }
            }

            #endregion

            List<TableModel> modelList = new List<TableModel>();

            #region Получение Comment, CaseType, Name -> Deviation

            foreach (string str in strList)
            {
                TableModel model = new TableModel();

                string currentValue = string.Empty;

                for (int i = 0; i < str.Length; i++)
                {
                    char ch = str[i];

                    char? nextChar = null;
                    if (i < str.Length - 1)
                        nextChar = str[i + 1];

                    #region NAME (Value)
                    
                    // если последняя буква
                    if (i == str.Length - 1)
                    {
                        if (ch != ' ')
                            currentValue += ch;

                        if (model.CaseType == "-")
                        {
                            model.CaseType = currentValue;

                            break;
                        }

                        if (!string.IsNullOrEmpty(currentValue))
                        {
                            #region DEVIATION

                            if (currentValue.Contains((char) 177))
                            {
                                int plusMinusIndex = currentValue.IndexOf((char) 177);
                
                                model.Deviation = currentValue.Substring(plusMinusIndex, currentValue.IndexOf('%') - plusMinusIndex + 1);
                            }
                            
                            #endregion
                            
                            #region NOMINAL
                                
                            // было решено привязываться к спецсимволам-разделителям, чтобы узнать номинал
                            
                            int? firstIndex = default;
                            int? secondIndex = default;

                            for (int u = currentValue.Length - 1; u > -1; u--)
                            {
                                char nCh = currentValue[u];

                                if (nCh == '_')
                                {
                                    if (secondIndex == null)
                                    {
                                        secondIndex = u;
                                    }
                                    else
                                    {
                                        firstIndex = u;
                                    }
                                }
                            }

                            if (secondIndex != null)
                            {
                                if (firstIndex != null)
                                {
                                    model.Nominal = currentValue.Substring(firstIndex.GetValueOrDefault() + 1,
                                        secondIndex.GetValueOrDefault() - firstIndex.GetValueOrDefault() - 1)
                                        .Replace(" ", string.Empty);
                                }
                                else
                                {
                                    model.Nominal = currentValue.Substring(secondIndex.GetValueOrDefault() + 1);
                                }
                                
                                model.Nominal = model.Nominal.Replace(" ", string.Empty);
                            }

                            #endregion

                            model.Name = currentValue.Replace('_', ' ');
                            
                            currentValue = string.Empty;
                        }
                    }

                    #endregion

                    // если череда пробелов - пропускаем
                    if (ch == ' ' && (nextChar == null || nextChar == ' '))
                        continue;

                    char? prevChar = null;
                    if (i > 0)
                        prevChar = str[i - 1];

                    // если не пробел или если по соседству с пробелом есть непробелы - пишем значение
                    if (ch != ' ' ||
                        ch == ' ' && prevChar != null && !char.IsWhiteSpace(prevChar.GetValueOrDefault()) &&
                        nextChar != null && !char.IsWhiteSpace(nextChar.GetValueOrDefault()))
                    {
                        currentValue += ch;

                        continue;
                    }

                    if (!string.IsNullOrEmpty(currentValue))
                    {
                        // Comment
                        if (model.Comment == "-")
                        {
                            model.Comment = currentValue.Replace(" ", string.Empty);

                            model.Quanity.ZeroZero = model.Quanity.ZeroOne = 1;

                            string componentTypeFromTxt = new string(model.Comment.Where(ch => !char.IsDigit(ch)).ToArray());

                            model.ComponentType = Dictionaries.ComponentTypes.GetValueOrDefault(componentTypeFromTxt, "-");
                        }
                        // CaseType
                        else if (model.CaseType == "-")
                        {
                            // заменяем C_**** на ****
                            if (currentValue[0] == 'C' && currentValue[1] == '_')
                                model.CaseType = currentValue.Remove(0, 2);
                            else
                                model.CaseType = currentValue;
                            
                            // подстраховка (убираем пробелы)
                            model.CaseType = model.CaseType.Replace(" ", string.Empty);
                        }

                        // чистим строку
                        currentValue = string.Empty;
                    }
                }

                modelList.Add(model);
            }

            #endregion

            List<TableModel> resultList = new List<TableModel>();

            #region ОБЪЕДИНЕНИЕ ОДИНАКОВЫХ КОМПОНЕНТОВ

            foreach (TableModel tm in modelList)
            {
                TableModel sameTM = resultList.FirstOrDefault(r =>
                                                                r.Name == tm.Name &&
                                                                r.CaseType == tm.CaseType &&
                                                                r.Deviation == tm.Deviation &&
                                                                r.Nominal == tm.Nominal);

                if (sameTM == null)
                {
                    resultList.Add(tm);
                }
                else if (!sameTM.Comment.Contains(tm.Comment))
                {
                    sameTM.Comment += $", {tm.Comment}";

                    sameTM.Quanity.ZeroZero = sameTM.Quanity.ZeroOne += 1;
                }
            }

            #endregion

            return resultList.OrderBy(r => r.ComponentType).ToArray();
        }
    }
}