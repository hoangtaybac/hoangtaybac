# mt2mml_worker.rb
# ------------------------------------------------------------
# Persistent Ruby worker for OLE(.bin) -> MathML conversion.
#
# Reads JSON lines from STDIN:
#   {"id":"...","b64":"..."}
#
# Writes JSON lines to STDOUT:
#   {"id":"...","ok":true,"mathml":"<math>...</math>"}
#   {"id":"...","ok":false,"err":"..."}
#
# Requires mt2mml_v2.rb (which must define Mt2Mml.convert_path).
# ------------------------------------------------------------

require "json"
require "base64"
require "tmpdir"
require "securerandom"

begin
  require File.join(Dir.pwd, "mt2mml_v2.rb")
rescue LoadError, StandardError => e
  # If it fails, we'll still run but every job will return an error.
  @__load_error = e
end

STDOUT.sync = true
STDERR.sync = true

def convert_b64_to_mathml(b64)
  raise @__load_error if defined?(@__load_error) && @__load_error

  bin = Base64.decode64(b64.to_s)
  Dir.mktmpdir("mt-ole-") do |dir|
    p = File.join(dir, "oleObject.bin")
    File.binwrite(p, bin)
    return Mt2Mml.convert_path(p).to_s
  end
end

ARGF.each_line do |line|
  line = line.to_s.strip
  next if line.empty?

  begin
    req = JSON.parse(line)
    id  = req["id"].to_s
    b64 = req["b64"].to_s

    mathml = convert_b64_to_mathml(b64)

    mathml = mathml.to_s.strip
    if mathml.empty? || !mathml.start_with?("<")
      puts JSON.generate({ id: id, ok: false, err: "empty_mathml" })
    else
      puts JSON.generate({ id: id, ok: true, mathml: mathml })
    end
  rescue => e
    id = ""
    begin
      id = (JSON.parse(line)["id"] rescue "").to_s
    rescue
      id = ""
    end
    puts JSON.generate({ id: id, ok: false, err: "#{e.class}: #{e.message}" })
  end
end
