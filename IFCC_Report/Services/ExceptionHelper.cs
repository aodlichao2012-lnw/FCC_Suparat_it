using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Bunker.Web.Services
{
    public static class ExceptionHelper
    {
        private static Queue<Exception> exceptions;

        public static void AddException(Exception ex)
        {
            if (exceptions == null)
            {
                exceptions = new Queue<Exception>();
            }
            exceptions.Enqueue(ex);
        }

        public static Exception GetException()
        {
            if (exceptions == null)
            {
                exceptions = new Queue<Exception>();
            }
            if (exceptions.Any())
            {
                return exceptions.Dequeue();
            }
            else
            {
                return null;
            }
        }

        public static void Clear()
        {
            exceptions.Clear();
        }
    }
}