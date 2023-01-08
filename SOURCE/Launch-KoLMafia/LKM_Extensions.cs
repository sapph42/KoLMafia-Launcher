using HtmlAgilityPack;
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace Launch_KoLMafia {
    internal static class LKM_Extensions {
        [return: MaybeNull]
        public static string? Hash(this FileInfo file,
                                    HashAlgorithm cryptoService) {
            if (!file.Exists) return null;
            StringBuilder builder = new();
            try {
				using FileStream fileStream = file.Open(FileMode.Open);
				fileStream.Position = 0;
				byte[] bytes = cryptoService.ComputeHash(fileStream);
				foreach (byte b in bytes) {
					builder.Append(b.ToString("x2"));
				}
			} catch (System.IO.IOException) {
                return null;
            }          
            return builder.ToString().ToLower();
        }
        [return: MaybeNull]
        public static string? GetFirstMatchingDescendent(this HtmlNode? node,
                                                        string ElementType,
                                                        [StringSyntax(StringSyntaxAttribute.Regex)] string Pattern,
                                                        bool MatchOnly = false) {
            if (node is null) return null;
            foreach (HtmlNode dNode in node.Descendants(ElementType)) {
                if (dNode.NodeType == HtmlNodeType.Element && Regex.IsMatch(dNode.InnerHtml, Pattern, RegexOptions.IgnoreCase)) {
                    if (MatchOnly) {
                        return Regex.Match(dNode.InnerHtml, Pattern, RegexOptions.IgnoreCase).Captures[0].Value;
                    }
                    return dNode.InnerHtml;
                }
            }
            return null;
        }
        [return: MaybeNull]
        public static HtmlNode? GetUriBody(this Uri uri) {
            HtmlWeb web = new();
            HtmlAgilityPack.HtmlDocument htmlDoc = web.Load(uri);
            return htmlDoc.DocumentNode.SelectSingleNode("//body");
        }

        public static bool IsGreater(this Version leftVersion, Version rightVersion) {
            if (leftVersion is null) return false;
            if (rightVersion is null) return false;
            if (leftVersion.Major > rightVersion.Major) {
                return true;
            }
            if (leftVersion.Minor > rightVersion.Minor) {
                return true;
            }
            if (leftVersion.Build > rightVersion.Build) {
                return true;
            }
            if (leftVersion.Revision > rightVersion.Revision) {
                return true;
            }
            return false;
        }
    }
}
