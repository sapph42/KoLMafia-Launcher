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
            using (cryptoService) {
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
            }
            return builder.ToString().ToLower();
        }
        [return: MaybeNull]
        public static string? GetFirstMatchingDescendent(this HtmlNode? node,
                                                        string ElementType,
                                                        [StringSyntax(StringSyntaxAttribute.Regex)] string Pattern) {
            if (node is null) return null;
            foreach (HtmlNode dNode in node.Descendants(ElementType)) {
                if (dNode.NodeType == HtmlNodeType.Element && Regex.IsMatch(dNode.InnerHtml, Pattern, RegexOptions.IgnoreCase)) {
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
    }
}
