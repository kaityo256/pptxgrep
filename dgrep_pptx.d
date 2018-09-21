import std.stdio;
import std.file;
import std.regex;
import std.zip;
import std.string;
import std.algorithm;
import std.path;

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
      if(match(str, keyword)){
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
