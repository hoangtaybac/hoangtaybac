# mt_worker.rb — Persistent Ruby worker for MathType OLE(.bin) -> MathML
# Protocol: one JSON request per line on STDIN, one JSON response per line on STDOUT
# Request:  {"id":"...","path":"/tmp/ole.bin","mode":"fast"|"v2"|"auto"}
# Response: {"id":"...","ok":true,"mathml":"<math>...</math>","mode_used":"fast"|"v2"} or {"id":"...","ok":false,"error":"..."}

STDOUT.sync = true
STDERR.sync = true

require "json"

# Optional gems (keep worker alive even if something is missing)
$gem_mathml = false
$gem_mathtype = false
$has_rexml = false

begin
  require "mathtype_to_mathml"
  $gem_mathml = true
rescue LoadError => e
  STDERR.puts "Warning: mathtype_to_mathml gem not found: #{e.message}"
end

begin
  require "mathtype"
  $gem_mathtype = true
rescue LoadError => e
  STDERR.puts "Warning: mathtype gem not found: #{e.message}"
end

begin
  require "rexml/document"
  $has_rexml = true
rescue LoadError => e
  STDERR.puts "Warning: rexml not found: #{e.message}"
end

# Detect sqrt in MathML (covers common entity forms)
SQRT_MATHML_RE = /(msqrt|mroot|√|&#8730;|&#x221a;|&#x221A;|&radic;)/i

def need_v2_for_sqrt?(mathml)
  s = mathml.to_s
  return false if s.empty?
  return false unless SQRT_MATHML_RE.match?(s)

  # already good
  return false if s =~ /<msqrt\b|<mroot\b/i

  # explicit bad pattern <mo>√</mo>...
  return true if s =~ /<mo>\s*(?:√|&#8730;|&#x221a;|&#x221A;|&radic;)\s*<\/mo>/i

  # has sqrt but not msqrt/mroot => likely bad
  true
end

def xml_escape(s)
  s.to_s
   .gsub("&", "&amp;")
   .gsub("<", "&lt;")
   .gsub(">", "&gt;")
   .gsub('"', "&quot;")
   .gsub("'", "&apos;")
end

# ------------------------
# Custom MTEF XML -> MathML (subset, focused on tmROOT correctness)
# ------------------------
class MtefToMathml
  def initialize(mtef_xml)
    raise "rexml missing" unless $has_rexml
    @doc = REXML::Document.new(mtef_xml.to_s)
  end

  def convert
    mtef = find_mtef_node
    return "" unless mtef

    main_slot = nil
    found_full = false

    mtef.each_element do |el|
      if el.name == "full"
        found_full = true
      elsif el.name == "slot" && found_full
        main_slot = el
        break
      end
    end

    # fallback: some xml may not have <full> wrapper
    main_slot ||= mtef.elements["slot"]
    return "" unless main_slot

    content = convert_slot(main_slot)
    %Q(<math xmlns="http://www.w3.org/1998/Math/MathML"><mrow>#{content}</mrow></math>)
  end

  private

  def find_mtef_node
    r = @doc.root
    return nil unless r
    return r if r.name == "mtef"

    # try direct child
    m = r.elements["mtef"]
    return m if m

    # scan descendants (safe)
    @doc.elements.each("//*") do |el|
      return el if el.name == "mtef"
    end
    nil
  end

  def convert_slot(slot_el)
    return "" unless slot_el
    result = +""
    slot_el.each_element do |el|
      case el.name
      when "char"
        result << convert_char(el)
      when "tmpl"
        result << convert_template(el)
      when "slot"
        result << convert_slot(el)
      end
    end
    result
  end

  def convert_char(char_el)
    mt_code = char_el.elements["mt_code_value"]&.text
    return "" unless mt_code
    t = mt_code.strip
    code = t.start_with?("0x") ? t.to_i(16) : t.to_i

    # Skip control/private chars that break downstream
    return "" if code < 0x0020
    return "" if code == 0x007F
    return " " if code == 0x00A0
    return "" if (0x200B..0x200D).cover?(code)
    return "" if code == 0xFEFF
    return "" if (0xE000..0xF8FF).cover?(code)

    char = [code].pack("U")
    esc = xml_escape(char)

    # normalize star-like chars
    return "<mo>*</mo>" if code == 0x2217 || code == 0x22C6

    if char =~ /[0-9]/
      "<mn>#{esc}</mn>"
    elsif char =~ /[a-zA-Z]/
      "<mi>#{esc}</mi>"
    else
      "<mo>#{esc}</mo>"
    end
  end

  def convert_template(tmpl_el)
    selector  = tmpl_el.elements["selector"]&.text&.strip
    variation = tmpl_el.elements["variation"]&.text&.strip

    slots = []
    tmpl_el.each_element("slot") { |s| slots << s }

    case selector
    when "tmROOT"
      radicand_slot = slots.find { |s| s.elements["options"]&.text&.strip != "1" } || slots[0]
      index_slot    = slots.find { |s| s.elements["options"]&.text&.strip == "1" }

      radicand = convert_slot(radicand_slot)

      if variation == "tvROOT_SQ" || index_slot.nil? || slot_is_empty?(index_slot)
        "<msqrt><mrow>#{radicand}</mrow></msqrt>"
      else
        index = convert_slot(index_slot)
        "<mroot><mrow>#{radicand}</mrow><mrow>#{index}</mrow></mroot>"
      end

    when "tmFRACT"
      num = slots[0] ? convert_slot(slots[0]) : ""
      den = slots[1] ? convert_slot(slots[1]) : ""
      "<mfrac><mrow>#{num}</mrow><mrow>#{den}</mrow></mfrac>"

    when "tmSUP"
      base = slots[0] ? convert_slot(slots[0]) : ""
      exp  = slots[1] ? convert_slot(slots[1]) : ""
      "<msup><mrow>#{base}</mrow><mrow>#{exp}</mrow></msup>"

    when "tmSUB"
      base = slots[0] ? convert_slot(slots[0]) : ""
      sub  = slots[1] ? convert_slot(slots[1]) : ""
      "<msub><mrow>#{base}</mrow><mrow>#{sub}</mrow></msub>"

    when "tmSUBSUP"
      base = slots[0] ? convert_slot(slots[0]) : ""
      sub  = slots[1] ? convert_slot(slots[1]) : ""
      sup  = slots[2] ? convert_slot(slots[2]) : ""
      "<msubsup><mrow>#{base}</mrow><mrow>#{sub}</mrow><mrow>#{sup}</mrow></msubsup>"

    when "tmPAREN"
      inner = slots[0] ? convert_slot(slots[0]) : ""
      "<mfenced open=\"(\" close=\")\"><mrow>#{inner}</mrow></mfenced>"

    when "tmBRACK"
      inner = slots[0] ? convert_slot(slots[0]) : ""
      "<mfenced open=\"[\" close=\"]\"><mrow>#{inner}</mrow></mfenced>"

    when "tmBRACE"
      inner = slots[0] ? convert_slot(slots[0]) : ""
      "<mfenced open=\"{\" close=\"}\"><mrow>#{inner}</mrow></mfenced>"

    when "tmBAR", "tmOBAR"
      inner = slots[0] ? convert_slot(slots[0]) : ""
      "<mover><mrow>#{inner}</mrow><mo>¯</mo></mover>"

    when "tmVEC"
      inner = slots[0] ? convert_slot(slots[0]) : ""
      "<mover><mrow>#{inner}</mrow><mo>→</mo></mover>"

    when "tmHAT"
      inner = slots[0] ? convert_slot(slots[0]) : ""
      "<mover><mrow>#{inner}</mrow><mo>^</mo></mover>"

    when "tmLIM"
      if slots.length >= 1
        sub = convert_slot(slots[0])
        "<munder><mo>lim</mo><mrow>#{sub}</mrow></munder>"
      else
        "<mo>lim</mo>"
      end

    else
      slots.map { |s| convert_slot(s) }.join
    end
  end

  def slot_is_empty?(slot_el)
    return true unless slot_el
    slot_el.each_element do |el|
      return false if %w[char tmpl].include?(el.name)
    end
    true
  end
end

# ------------------------
# Converters
# ------------------------
def convert_fast(file_path)
  raise "File not found: #{file_path}" unless file_path && File.exist?(file_path)
  raise "mathtype_to_mathml gem missing" unless $gem_mathml

  converter = MathTypeToMathML::Converter.new(file_path)
  (converter.convert || "").to_s
end

def convert_v2(file_path)
  raise "File not found: #{file_path}" unless file_path && File.exist?(file_path)

  # If mathtype gem missing, fall back to fast (if available)
  unless $gem_mathtype
    return convert_fast(file_path) if $gem_mathml
    raise "Required gems missing: mathtype_to_mathml and/or mathtype"
  end

  mtef_xml = Mathtype::Converter.new(file_path).to_xml.to_s
  has_root = (mtef_xml =~ /<selector>\s*tmROOT\s*<\/selector>/i) ? true : false

  # Prefer custom when tmROOT exists (sqrt/mroot correctness)
  if has_root && $has_rexml
    begin
      mm = MtefToMathml.new(mtef_xml).convert.to_s
      return mm if mm =~ /<math\b/i && mm =~ /<(msqrt|mroot)\b/i
    rescue => e
      STDERR.puts "v2 custom converter error: #{e.class}: #{e.message}"
    end
  end

  # Otherwise try gem fast converter (often best for non-root templates)
  if $gem_mathml
    begin
      mm = convert_fast(file_path)
      return mm if mm =~ /<math\b/i
    rescue => e
      STDERR.puts "v2 gem converter error: #{e.class}: #{e.message}"
    end
  end

  # Final fallback: custom converter even if not root (only if rexml exists)
  if $has_rexml
    begin
      mm = MtefToMathml.new(mtef_xml).convert.to_s
      return mm if mm =~ /<math\b/i
    rescue => e
      STDERR.puts "v2 final custom converter error: #{e.class}: #{e.message}"
    end
  end

  ""
end

def convert_auto(file_path)
  # Strategy:
  # 1) Try fast first (if available)
  # 2) If fast result suggests broken sqrt -> run v2 and prefer v2 if it yields msqrt/mroot
  # 3) If fast not available or empty -> try v2
  fast_mm = ""
  if $gem_mathml
    begin
      fast_mm = convert_fast(file_path).to_s
      if !fast_mm.empty?
        if need_v2_for_sqrt?(fast_mm)
          v2_mm = convert_v2(file_path).to_s
          if v2_mm =~ /<math\b/i && v2_mm =~ /<(msqrt|mroot)\b/i
            return [v2_mm, "v2"]
          end
        end
        return [fast_mm, "fast"]
      end
    rescue => e
      STDERR.puts "auto fast error: #{e.class}: #{e.message}"
    end
  end

  begin
    v2_mm = convert_v2(file_path).to_s
    return [v2_mm, "v2"] unless v2_mm.empty?
  rescue => e
    STDERR.puts "auto v2 error: #{e.class}: #{e.message}"
  end

  ["", "fast"]
end

def respond(hash)
  puts JSON.generate(hash)
rescue => e
  puts %Q({"ok":false,"error":"JSON error: #{e.class}: #{e.message}"})
end

# ------------------------
# Main loop
# ------------------------
while (line = STDIN.gets)
  s = line.to_s.strip
  next if s.empty?

  req = nil
  begin
    req = JSON.parse(s)
  rescue => e
    respond({ ok: false, error: "Bad JSON: #{e.class}: #{e.message}" })
    next
  end

  id   = req["id"]
  pth  = req["path"]
  mode = (req["mode"] || "fast").to_s

  begin
    mathml = ""
    mode_used = "fast"

    if mode == "v2"
      mathml = convert_v2(pth)
      mode_used = "v2"
    elsif mode == "auto"
      mathml, mode_used = convert_auto(pth)
    else
      mathml = convert_fast(pth)
      mode_used = "fast"
    end

    mathml = mathml.to_s.strip
    ok = !mathml.empty? && mathml.start_with?("<") && (mathml =~ /<math\b/i)

    respond({ id: id, ok: ok, mathml: ok ? mathml : "", mode_used: mode_used })
  rescue => e
    respond({ id: id, ok: false, error: "#{e.class}: #{e.message}", mode_used: mode })
  end
end
