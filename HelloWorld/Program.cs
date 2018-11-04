using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Se.Url;

namespace HelloWorld
{
    class Program
    {
        static void Main(string[] args)
        {
            var hello = new HelloWorld("cn");
            hello.SayHello();
            hello.SayHello(5 as Object);
            hello.SayHello<HelloWorld>(new HelloWorld("en"));

            var cosmo = new HelloComso<Thread>(new Thread(() => { Console.WriteLine("doing nothing"); }));
            cosmo.SayHello();

            var urlBlder = new UrlBuilder("http://www.google.com?search=c#");
            Console.WriteLine($"{urlBlder.GetHost()}");
            Console.Write("Press any key to exit...");
            Console.Read();
        }
    }

    class HelloWorld
    {
        private string _lang;

        public HelloWorld()
        {
            _lang = "en";
        }
        public HelloWorld(string lang)
        {
            _lang = lang;
        }

        public string Language => _lang;

        public virtual void SayHello()
        {
            switch (Language)
            {
                case "cn":
                    Console.WriteLine("你好，世界！");
                    break;
                case "en":
                default:
                    Console.WriteLine("Hello, World!");
                    break;

            }
        }

        public virtual void SayHello<T>(T words) where T : class
        {
            Console.WriteLine($"{typeof(T).ToString()} : {words.ToString()}");
        }
    }

    class HelloComso<T> : HelloWorld where T : class
    {
        private T _words;

        public HelloComso(T words)
        {
            _words = words;
        }
        public override void SayHello()
        {
            base.SayHello(_words);
            if (_words is Thread)
            {
                ((_words) as Thread).Start();
                Thread.Sleep(100);
            }
        }
    }
}
