# mt2mml_v2.rb
require File.join(Dir.pwd, "mt2mml.rb")

module Mt2Mml
  # Mt2Mml.convert_path is already defined in mt2mml.rb
end

if __FILE__ == $0
  pth = ARGV[0]
  begin
    puts Mt2Mml.convert_path(pth)
  rescue => e
    warn e.message
    exit 1
  end
end
