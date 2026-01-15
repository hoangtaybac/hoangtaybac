#!/usr/bin/env ruby
# Convert MathType OLE (.bin) to MathML and print to stdout.
# Requires gem: mathtype_to_mathml_plus
# Usage: ruby mt2mml.rb /path/to/oleObject1.bin

begin
  require 'mathtype_to_mathml_plus'
rescue LoadError => e
  warn "Missing gem mathtype_to_mathml_plus. Install: gem install mathtype_to_mathml_plus"
  raise
end

in_path = ARGV[0]
if in_path.nil? || in_path.strip.empty?
  warn "Usage: ruby mt2mml.rb /path/to/oleObject.bin"
  exit 2
end

begin
  mathml = MathTypeToMathMLPlus::Converter.new(in_path).convert
  puts mathml.to_s
rescue => e
  warn e.message
  exit 1
end
