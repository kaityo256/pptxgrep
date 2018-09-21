import std.stdio;
import std.file;
import std.regex;
import std.zip;
import std.string;
import std.algorithm;
import std.path;
import std.xml;
import std.conv;

dstring extractText(string xmltext)
{
  dstring dxml = xmltext.to!dstring();
  dstring text;
  while(findSkip(dxml, "<a:t>")){
    auto e = indexOf(dxml,"</a:t>");
    text ~= dxml[0..e];
  }
  return text;
}

void search(string keyword, string filename)
{
  auto zip = new ZipArchive(read(filename));
  foreach (name, am; zip.directory)
  {
    foreach(m; match(name, r"ppt/slides/slide([0-9]+).xml$"))
    {
      zip.expand(am);
      auto slidenum = m.captures[1];
      char *cstr = cast(char*)am.expandedData;
      auto len = am.expandedData.length;
      string str = cast(string) cstr[0..len];
      // This is XML version. It is too slow.
      /*
      auto xml = new DocumentParser(str);
      dstring text;
      xml.onText = (string s)
      {
        text ~= s.to!dstring;
      };
      xml.parse();
      */
      dstring text = extractText(str);
      if(text.indexOf(keyword) !=-1)
      {
        auto rname = relativePath(filename);
        writefln("Found \"%s\" in %s at slide %s",keyword,rname, slidenum);
      }
    } 
  }
}

void main(string[] args)
{
  if(args.length <2)
  {
    writeln("Usage:");
    writeln("  dgrep_pptx keyword");
    return;
  }
  auto keyword = args[1];
  auto cwd = std.file.getcwd();
  auto d = dirEntries(cwd,"*.pptx",SpanMode.depth);
  string [] files;
  foreach(string filename; d){
    files ~= filename;
  }
  files.sort!();
  foreach(string filename; files){
    search(keyword, filename);
  }
}
