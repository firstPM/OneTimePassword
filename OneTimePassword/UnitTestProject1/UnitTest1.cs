using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using System.Diagnostics;
using Cryptolens.OneTimePassword;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var secret = OneTimePassword.CreateSharedSecret();

            Debug.WriteLine("CounterBasedPassword:" + OneTimePassword.CounterBasedPassword(secret, 123));
            Debug.WriteLine("TimeBasedPassword:" + OneTimePassword.TimeBasedPassword(new byte[] { 123, 123 }));

            Debug.WriteLine(secret);
        }

        [TestMethod]
        public void MyTestMethod()
        {
            var secret = OneTimePassword.CreateSharedSecret();
            Debug.WriteLine("SharedSecretToString:" + OneTimePassword.SharedSecretToString(secret));

            Debug.WriteLine("TimeBasedPassword:" + OneTimePassword.TimeBasedPassword(secret));
            //OneTimePassword.SharedSecretToString(secret) 
            //TCAAZESBWKU3YQZ7OYG7JTM6QZYX2HNM
            Debug.WriteLine("AppUrl:"+OneTimePassword.GetAuthenticatorAppUrl("type secret key or use OneTimePassword.SharedSecretToString(secret) ", "long.ming@compass.com", "auth"));
        }
    }
}
