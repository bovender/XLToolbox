using System;
using System.Reflection;
using Bovender.Versioning;
using NUnit.Framework;

namespace XLToolbox.Test
{
    [TestFixture]
    public class VersioningTest
    {
        [Test]
        [TestCase("0.1.2", 0, 1, 2, "", "")]
        [TestCase("10.20.30-0.0.1", 10, 20, 30, "0.0.1", "")]
        [TestCase("0.1.2-alpha.1", 0, 1, 2, "alpha.1", "")]
        [TestCase("0.1.2-0.0.1+githash", 0, 1, 2, "0.0.1", "githash")]
        [TestCase("0.1.2+githash", 0, 1, 2, "", "githash")]
        public void ParseSemanticVersion(string version, int major, int minor, int patch,
            string preRelease, string build)
        {
            SemanticVersion semVer = new SemanticVersion(version);
            Assert.AreEqual(major, semVer.Major, "Major version does not match");
            Assert.AreEqual(minor, semVer.Minor, "Minor version does not match");
            Assert.AreEqual(patch, semVer.Patch, "Patch number does not match");
            Assert.AreEqual(build, semVer.Build, "Build information does not match");
        }

        [Test]
        public void GetCurrentVersion()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            SemanticVersion v = SemanticVersion.CurrentVersion(assembly);
            Assert.AreEqual(7, v.Major);
        }

        [Test]
        [ExpectedException(typeof(InvalidVersionStringException))]
        public void InvalidVersionThrowsError()
        {
            SemanticVersion v = new SemanticVersion("2.0");
        }

        [Test]
        [TestCase("1.0.0", "1.0.1")]
        [TestCase("1.0.0", "1.1.0")]
        [TestCase("1.0.1", "2.0.0")]
        [TestCase("1.1.1", "2.0.0")]
        [TestCase("1.1.1-0.1.2", "1.1.1")]
        [TestCase("1.1.1-alpha.10", "1.1.1")]
        [TestCase("1.1.1-beta.2", "1.1.1")]
        [TestCase("1.1.1-alpha.2", "1.1.1-beta.1")]
        [TestCase("1.1.1-rc.1", "1.1.1-rc.2")]
        public void CompareVersions(string lower, string higher)
        {
            SemanticVersion lowerVersion = new SemanticVersion(lower);
            SemanticVersion higherVersion = new SemanticVersion(higher);
            Assert.Greater(higherVersion, lowerVersion);
            Assert.True(lowerVersion < higherVersion);
            Assert.True(higherVersion > lowerVersion);
        }
    }
}
