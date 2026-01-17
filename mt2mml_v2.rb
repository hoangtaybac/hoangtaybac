#!/usr/bin/env ruby
# mt2mml_v2.rb
# Usage: ruby mt2mml_v2.rb /path/to/oleObject.bin

in_path = ARGV[0]
abort("") unless in_path && File.exist?(in_path)

def try_convert(const_name, file_path)
  Object.const_get(const_name)::Converter.new(file_path).convert
rescue NameError
  nil
end

begin
  begin
    require "mathtype_to_mathml_plus"
    mml = try_convert("MathTypeToMathMLPlus", in_path)
    if mml && !mml.to_s.strip.empty?
      puts mml.to_s.strip
      exit 0
    end
  rescue LoadError
    # ignore -> fallback
  end

  require "mathtype_to_mathml"
  mml = try_convert("MathTypeToMathML", in_path)
  puts (mml ? mml.to_s.strip : "")
  exit 0
rescue => e
  warn e.message
  puts ""
  exit 0
end
