# mt2mml.rb - MathType OLE(.bin) -> MathML
require "mathtype_to_mathml"

module Mt2Mml
  def self.convert_path(path)
    converter = MathTypeToMathML::Converter.new(path)
    converter.convert.to_s
  end
end

# CLI compatible
if __FILE__ == $0
  path = ARGV[0]
  abort "usage: ruby mt2mml.rb <oleObject*.bin>" unless path && File.exist?(path)

  begin
    puts Mt2Mml.convert_path(path)
  rescue => e
    STDERR.puts "Error: #{e.message}"
    exit 1
  end
end
