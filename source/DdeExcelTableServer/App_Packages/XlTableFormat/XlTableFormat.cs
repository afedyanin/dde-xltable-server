#region Copyright (c) 2014 Atif Aziz. All rights reserved.
//
// Copyright (c) 2014 Atif Aziz. All rights reserved.
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//    http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
#endregion

#region Imports

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

#endregion

// ReSharper disable once PartialTypeWithSinglePart
partial interface IXlTableDataFactory<out T>
{
    T Blank { get; }
    T Skip  { get; }
    T Table (int rows, int cols);
    T Float (double value);
    T String(string value);
    T Bool  (bool value);
    T Error (int value);
    T Int   (int value);
}

// ReSharper disable once PartialTypeWithSinglePart
sealed partial class XlTableDataFactory<T>
{
    public T Blank { get; set; }
    public T Skip { get; set; }
    public Func<int, int, T> Table { get; set; }
    public Func<double, T> Float { get; set; }
    public Func<string, T> String { get; set; }
    public Func<bool, T> Bool { get; set; }
    public Func<int, T> Error { get; set; }
    public Func<int, T> Int { get; set; }

    public IXlTableDataFactory<T> Bind() { return new Closure(this); }

    sealed class Closure : IXlTableDataFactory<T>
    {
        readonly T _blank;
        readonly T _skip;
        readonly Func<int, int, T> _table;
        readonly Func<double, T> _float;
        readonly Func<string, T> _string;
        readonly Func<bool, T> _bool;
        readonly Func<int, T> _error;
        readonly Func<int, T> _int;

        public Closure(XlTableDataFactory<T> factory)
        {
            _blank  = factory.Blank;
            _skip   = factory.Skip;
            _table  = factory.Table  ?? delegate { return default(T); };
            _float  = factory.Float  ?? delegate { return default(T); };
            _string = factory.String ?? delegate { return default(T); };
            _bool   = factory.Bool   ?? delegate { return default(T); };
            _error  = factory.Error  ?? delegate { return default(T); };
            _int    = factory.Int    ?? delegate { return default(T); };
        }

        T IXlTableDataFactory<T>.Blank                     { get { return _blank; }     }
        T IXlTableDataFactory<T>.Skip                      { get { return _skip;  }     }
        T IXlTableDataFactory<T>.Table(int rows, int cols) { return _table(rows, cols); }
        T IXlTableDataFactory<T>.Float(double value)       { return _float(value);      }
        T IXlTableDataFactory<T>.String(string value)      { return _string(value);     }
        T IXlTableDataFactory<T>.Bool(bool value)          { return _bool(value);       }
        T IXlTableDataFactory<T>.Error(int value)          { return _error(value);      }
        T IXlTableDataFactory<T>.Int(int value)            { return _int(value);        }
    }

}

// ReSharper disable once PartialTypeWithSinglePart
static partial class XlTableFormat
{
    public static readonly IXlTableDataFactory<object> DefaultDataFactory = new XlTableDataFactory<object>
    {
        Blank   = null,
        Skip    = Missing.Value,
        Table   = (rows, cols) => new[] { rows, cols },
        Float   = v => v,
        String  = v => v,
        Bool    = v => v,
        Error   = v => new ErrorWrapper(v),
        Int     = v => v,
    }
    .Bind();

    public static T Read<T>(byte[] data, Func<int, int, object[], T> resultor)
    {
        using (var ms = new MemoryStream(data))
            return Read(ms, resultor);
    }

    public static T Read<T>(Stream stream, Func<int, int, object[], T> resultor)
    {
        if (stream == null) throw new ArgumentNullException("stream");
        if (resultor == null) throw new ArgumentNullException("resultor");

        using (var e = Read(stream, DefaultDataFactory))
        {
            if (!e.MoveNext()) throw new FormatException();
            var size = (int[]) e.Current;
            var rows = size[0];
            var cols = size[1];
            var cells = new object[rows * cols];
            for (var i = 0; e.MoveNext(); i++)
                cells[i] = e.Current;
            return resultor(rows, cols, cells);
        }
    }

    public static IEnumerable<object> Read(byte[] data)
    {
        return Read(data, DefaultDataFactory);
    }

    public static IEnumerable<T> Read<T>(byte[] data, IXlTableDataFactory<T> factory)
    {
        if (data == null) throw new ArgumentNullException("data");
        if (factory == null) throw new ArgumentNullException("factory");
        return ReadImpl(data, factory);
    }
    
    static IEnumerable<T> ReadImpl<T>(byte[] data, IXlTableDataFactory<T> factory)
    {
        using (var ms = new MemoryStream(data))
        using (var e = Read(ms, factory))
        while (e.MoveNext())
            yield return e.Current;
    }

    public static IEnumerator<object> Read(Stream stream)
    {
        return Read(stream, DefaultDataFactory);
    }

    public static IEnumerator<T> Read<T>(Stream stream, IXlTableDataFactory<T> factory)
    {
        if (stream == null) throw new ArgumentNullException("stream");
        if (factory == null) throw new ArgumentNullException("factory");
        return ReadImpl(stream, factory);
    }
    
    static IEnumerator<T> ReadImpl<T>(Stream stream, IXlTableDataFactory<T> factory)
    {
        using (var reader = new BinaryReader(stream))
        {
            if (XlTableDataType.Table != (XlTableDataType) reader.ReadUInt16()) 
                throw new FormatException();
            var size = reader.ReadUInt16();
            if (size != 4) 
                throw new FormatException();
            var rows = reader.ReadUInt16();
            var cols = reader.ReadUInt16();
            yield return factory.Table(rows, cols);
            var cells = rows * cols;
            while (cells > 0)
            {
                var type = (XlTableDataType) reader.ReadUInt16();
                size = reader.ReadUInt16();
                if (type == XlTableDataType.String)
                {
                    while (size > 0)
                    {
                        var str = Encoding.Default.GetString(reader.ReadBytes(reader.ReadByte()));
                        yield return factory.String(str);
                        cells--;
                        size -= (ushort) (1 + checked((byte) str.Length));
                    }
                }
                else
                {
                    int count;
                    Func<BinaryReader, IXlTableDataFactory<T>, T> rf;
                    switch (type)
                    {
                        case XlTableDataType.Float:
                            if (size % 8 != 0) throw new FormatException();
                            count = size / 8;
                            rf = (r, f) => f.Float(r.ReadDouble());
                            break;
                        case XlTableDataType.Skip:
                            if (size != 2) throw new FormatException();
                            count = reader.ReadUInt16();
                            rf = (r, f) => f.Skip;
                            break;
                        case XlTableDataType.Blank:
                            if (size != 2) throw new FormatException();
                            count = reader.ReadUInt16();
                            rf = (r, f) => f.Blank;
                            break;
                        case XlTableDataType.Error:
                            if (size % 2 != 0) throw new FormatException();
                            count = size / 2;
                            rf = (r, f) => f.Error(r.ReadUInt16());
                            break;
                        case XlTableDataType.Bool:
                            if (size % 2 != 0) throw new FormatException();
                            count = size / 2;
                            rf = (r, f) => f.Bool(r.ReadUInt16() != 0);
                            break;
                        case XlTableDataType.Int:
                            if (size % 2 != 0) throw new FormatException();
                            count = size / 2;
                            rf = (r, f) => f.Int(r.ReadUInt16());
                            break;
                        default: throw new FormatException();
                    }

                    for (var j = 0; j < count; j++)
                        yield return rf(reader, factory);
                    
                    cells -= count;
                }
            }
        }
    }

    [Serializable]
    enum XlTableDataType
    {
        Table  = 16,
        Float  = 1,
        String = 2,
        Bool   = 3,
        Error  = 4,
        Blank  = 5,
        Int    = 6,
        Skip   = 7,
    }
}
