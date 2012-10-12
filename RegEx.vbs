using Microsoft.SqlServer.Server;
using System.Data.SqlTypes;
// namespace to work with regular expressions
using System.Text.RegularExpressions;
public class cls_RegularExpressions
{
  [SqlFunction]
  public static SqlString ReplaceMatch(
    SqlString InputString,
    SqlString MatchPattern,
    SqlString ReplacementPattern)
  {
    try
    {
      // input parameters must not be NULL
      if (!InputString.IsNull &&
          !MatchPattern.IsNull &&
          !ReplacementPattern.IsNull)
      {
      // check for first pattern match
      if (Regex.IsMatch(InputString.Value,
                        MatchPattern.Value))
        // match found, replace using second pattern and return result
        return Regex.Replace(InputString.Value,
                            MatchPattern.Value,
                            ReplacementPattern.Value);
      else
        // match not found, return NULL
        return SqlString.Null;
      }
      else
        // if any input paramater is NULL, return NULL
        return SqlString.Null;
    }
    catch
    {
      // on any error, return NULL
      return SqlString.Null;
    }
  }
};