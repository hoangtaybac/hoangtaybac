# mt2mml_v2.rb
# ------------------------------------------------------------
# Compatibility "v2" wrapper that exposes a stable API:
#
#   Mt2Mml.convert_path(path) -> MathML string
#
# This file tries (in order):
#  1) Require and call a converter API from ./mt2mml.rb (no subprocess)
#  2) If no API is found, fallback to executing: ruby mt2mml.rb <path>
#     and capturing stdout (works immediately, but slower).
#
# You can later wire this to your real fast converter by replacing
# Mt2Mml.convert_path implementation to call your internal functions.
# ------------------------------------------------------------

require "json"
require "open3"

module Mt2Mml
  # Try to load legacy converter script (if present) as a library.
  # This may define constants/methods we can call directly.
  def self._try_require_legacy!
    @legacy_loaded ||= begin
      legacy = File.join(Dir.pwd, "mt2mml.rb")
      if File.exist?(legacy)
        require legacy
        true
      else
        false
      end
    rescue LoadError, StandardError
      false
    end
  end

  # Try common converter entrypoints without spawning a new Ruby.
  def self._convert_via_loaded_api(path)
    # 1) MT2MML.convert_file / convert_path
    if defined?(MT2MML)
      return MT2MML.convert_file(path) if MT2MML.respond_to?(:convert_file)
      return MT2MML.convert_path(path) if MT2MML.respond_to?(:convert_path)
      return MT2MML.convert(path)      if MT2MML.respond_to?(:convert)
    end

    # 2) Mathtype2Mml / MathType2Mml classes/modules
    %i[Mathtype2Mml MathType2Mml Mathtype2MML MathType2MML].each do |c|
      next unless Object.const_defined?(c)
      obj = Object.const_get(c)
      return obj.convert_file(path) if obj.respond_to?(:convert_file)
      return obj.convert_path(path) if obj.respond_to?(:convert_path)
      return obj.convert(path)      if obj.respond_to?(:convert)
    end

    # 3) Converter module/class named Mt2Mml (older versions)
    if defined?(::Mt2Mml) && ::Mt2Mml.respond_to?(:convert_file)
      return ::Mt2Mml.convert_file(path)
    end

    nil
  end

  # Fallback: execute mt2mml.rb as a CLI and capture stdout
  def self._convert_via_cli(path)
    legacy = File.join(Dir.pwd, "mt2mml.rb")
    raise "mt2mml.rb not found in #{Dir.pwd}" unless File.exist?(legacy)

    stdout, stderr, status = Open3.capture3("ruby", legacy, path.to_s)
    raise "mt2mml.rb failed: #{stderr.to_s.strip}" unless status.success?

    out = stdout.to_s.strip
    # Some versions output JSON {mathml: "..."}; accept both.
    if out.start_with?("{")
      begin
        j = JSON.parse(out)
        out = (j["mathml"] || j["mml"] || j["out"] || "").to_s.strip
      rescue JSON::ParserError
        # keep as-is
      end
    end
    out
  end

  # Public API used by mt2mml_worker.rb
  def self.convert_path(path)
    _try_require_legacy!

    out = _convert_via_loaded_api(path)
    out = _convert_via_cli(path) if out.nil? || out.to_s.strip.empty?

    out = out.to_s.strip
    return "" if out.empty?
    return out if out.start_with?("<")

    # Sometimes the converter prints extra lines; try to extract first MathML tag.
    if (m = out.match(/(<math[\s\S]*<\/math>)/i))
      return m[1]
    end
    out
  end
end

# Keep CLI behavior too
if __FILE__ == $0
  pth = ARGV[0]
  begin
    puts Mt2Mml.convert_path(pth)
  rescue => e
    warn e.message
    exit 1
  end
end
