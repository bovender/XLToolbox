using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

namespace XLToolbox.Version
{
    public enum Prerelease
    {
        Numeric = 1,
        Alpha = 2,
        Beta = 3,
        RC = 4,
        None = 5
    }

    /// <summary>
    /// Class that handles semantic versioning.
    /// </summary>
    public class SemanticVersion : Object, IComparable
    {
        public int Major { get; set; }
        public int Minor { get; set; }
        public int Patch { get; set; }
        public Prerelease Prerelease { get; set; }
        public int PreMajor { get; set; }
        public int PreMinor { get; set; }
        public int PrePatch { get; set; }
        public string Build { get; set; }
        private string _version;

        /// <summary>
        /// Instantiates the class from a given version string.
        /// </summary>
        /// <param name="version">String that complies with semantic versioning rules.</param>
        public SemanticVersion(string version)
        {
            ParseString(version);
        }

        /// <summary>
        /// Factory method that creates an instance of the Version class with
        /// the current version information.
        /// </summary>
        /// <returns>Instance of Version</returns>
        public static SemanticVersion CurrentVersion()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream stream = assembly.GetManifestResourceStream("XLToolbox.Version.VERSION");
            StreamReader text = new StreamReader(stream);
            return new SemanticVersion(text.ReadLine());
        }

        /// <summary>
        /// Returns the full version string.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return _version;
        }

        public static bool operator <(SemanticVersion lower, SemanticVersion higher)
        {
            return (lower.CompareTo(higher) < 0);
        }

        public static bool operator >(SemanticVersion higher, SemanticVersion lower)
        {
            return (lower.CompareTo(higher) < 0);
        }

        public static bool operator <=(SemanticVersion lower, SemanticVersion higher)
        {
            return (lower.CompareTo(higher) <= 0);
        }

        public static bool operator >=(SemanticVersion higher, SemanticVersion lower)
        {
            return (lower.CompareTo(higher) <= 0);
        }

        public static bool operator ==(SemanticVersion v1, SemanticVersion v2)
        {
            return (v1.Equals(v2));
        }

        public static bool operator !=(SemanticVersion v1, SemanticVersion v2)
        {
            return (!v1.Equals(v2));
        }

        /// <summary>
        /// Parses a string that complies with semantic versioning, V. 2.
        /// </summary>
        /// <param name="s">Semantic version string.</param>
        protected void ParseString(string s)
        {
            Regex r = new Regex(
                @"(?<major>\d+)\.(?<minor>\d+)\.(?<patch>\d+)" +
                @"(-(?<pre>((?<preMajor>\d+)\.(?<preMinor>\d+)\.|"+
                @"((?<alpha>alpha)|(?<beta>beta)|(?<rc>rc))\.)(?<prePatch>\d+)))?" +
                @"(\+(?<build>[a-zA-Z0-9]+))?");
            Match m = r.Match(s);

            if (!m.Success)
            {
                throw new InvalidVersionStringException(s);
            };

            _version = s;
            Major = Convert.ToInt32(m.Groups["major"].Value);
            Minor = Convert.ToInt32(m.Groups["minor"].Value);
            Patch = Convert.ToInt32(m.Groups["patch"].Value);

            if (m.Groups["pre"].Success)
            {
                if (m.Groups["alpha"].Success)
                {
                    Prerelease = Version.Prerelease.Alpha;
                }
                else if (m.Groups["beta"].Success)
                {
                    Prerelease = Version.Prerelease.Beta;
                }
                else if (m.Groups["rc"].Success)
                {
                    Prerelease = Version.Prerelease.RC;
                }
                else
                {
                    Prerelease = Version.Prerelease.Numeric;
                    PreMajor = Convert.ToInt32(m.Groups["preMajor"].Value);
                    PreMinor = Convert.ToInt32(m.Groups["preMinor"].Value);
                }
            }
            else
            {
                Prerelease = Version.Prerelease.None;
            }
            if (m.Groups["prePatch"].Success)
            {
                PrePatch = Convert.ToInt32(m.Groups["prePatch"].Value);
            };

            Build = m.Groups["build"].Value;
        }

        public int CompareTo(object obj)
        {
            SemanticVersion other = obj as SemanticVersion;
            if (this.Major < other.Major)
            {
                return -1;
            }
            else if (this.Major > other.Major)
            {
                return 1;
            }
            else // both major versions are the same, compare minor version
            {
                if (this.Minor < other.Minor)
                {
                    return -1;
                }
                else if (this.Minor > other.Minor)
                {
                    return 1;
                }
                else // major and minor are same, compare patch
                {
                    if (this.Patch < other.Patch)
                    {
                        return -1;
                    }
                    else if (this.Patch > other.Patch)
                    {
                        return 1;
                    }
                    else // major, minor, and path are same, compare pre-release
                    {
                        if (this.Prerelease < other.Prerelease)
                        {
                            return -1;
                        }
                        else if (this.Prerelease > other.Prerelease)
                        {
                            return 1;
                        }
                        else // prerelease type is same (alpha/beta/etc.)
                        {
                            if (this.Prerelease == Version.Prerelease.Numeric)
                            {
                                if (this.PreMajor < other.PreMajor)
                                {
                                    return -1;
                                }
                                else if (this.PreMajor > other.PreMajor)
                                {
                                    return 1;
                                }
                                else
                                {
                                    if (this.PreMinor < other.PreMinor)
                                    {
                                        return -1;
                                    }
                                    else if (this.PreMinor > other.PreMinor)
                                    {
                                        return 1;
                                    }
                                    else
                                    {
                                        if (this.PrePatch < other.PrePatch)
                                        {
                                            return -1;
                                        }
                                        else if (this.PrePatch> other.PrePatch)
                                        {
                                            return 1;
                                        }
                                        else
                                        {
                                            return 0;
                                        }
                                    }
                                }
                            }
                            else // prerelease type same, not numeric
                            {
                                if (this.PrePatch < other.PrePatch)
                                {
                                    return -1;
                                }
                                else if (this.PrePatch > other.PrePatch)
                                {
                                    return 1;
                                }
                                else
                                {
                                    return 0;
                                }
                            }
                        }
                    }
                }
            }
        }

        public override bool Equals(object obj)
        {
            if (obj != null)
            {
                return (this.CompareTo(obj) == 0);
            }
            else
            {
                return false;
            }
        }

        public override int GetHashCode()
        {
            return _version.GetHashCode();
        }
    }
}
