# mt2mml_batch.rb - Batch converter for multiple OLE files
# Usage: ruby mt2mml_batch.rb file1.bin file2.bin file3.bin ...
# Output: JSON array of results

require 'json'

begin
  require 'mathtype'
rescue LoadError => e
  STDERR.puts "Warning: mathtype gem not found: #{e.message}"
end

begin
  require 'mathtype_to_mathml'
  GEM_AVAILABLE = true
rescue LoadError => e
  STDERR.puts "Warning: mathtype_to_mathml gem not found: #{e.message}"
  GEM_AVAILABLE = false
end

begin
  require 'rexml/document'
rescue LoadError => e
  STDERR.puts "Warning: rexml not found: #{e.message}"
end

# ============================================================
# Custom MTEF XML to MathML converter (same as mt2mml_v2.rb)
# ============================================================
class MtefToMathml
  def initialize(mtef_xml)
    @doc = REXML::Document.new(mtef_xml)
  end

  def convert
    mtef = @doc.root.elements['mtef']
    return nil unless mtef
    
    main_slot = nil
    found_full = false
    mtef.each_element do |el|
      if el.name == 'full'
        found_full = true
      elsif el.name == 'slot' && found_full
        main_slot = el
        break
      end
    end
    
    return nil unless main_slot
    
    content = convert_slot(main_slot)
    "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow>#{content}</mrow></math>"
  end

  private

  def convert_slot(slot_el)
    return "" unless slot_el
    
    result = ""
    slot_el.each_element do |el|
      case el.name
      when 'char'
        result += convert_char(el)
      when 'tmpl'
        result += convert_template(el)
      when 'slot'
        result += convert_slot(el)
      end
    end
    result
  end

  def convert_char(char_el)
    mt_code = char_el.elements['mt_code_value']&.text
    return "" unless mt_code
    
    code = mt_code.strip.start_with?('0x') ? mt_code.strip.to_i(16) : mt_code.strip.to_i
    
    return "" if code < 0x0020
    return "" if code == 0x007F
    return " " if code == 0x00A0
    return "" if code >= 0x200B && code <= 0x200D
    return "" if code == 0xFEFF
    return "" if code >= 0xE000 && code <= 0xF8FF
    
    char = [code].pack('U')
    
    return "<mo>*</mo>" if code == 0x2217 || code == 0x22C6
    
    if char =~ /[0-9]/
      "<mn>#{char}</mn>"
    elsif char =~ /[a-zA-Z]/
      "<mi>#{char}</mi>"
    else
      "<mo>#{char}</mo>"
    end
  end

  def convert_template(tmpl_el)
    selector = tmpl_el.elements['selector']&.text&.strip
    variation = tmpl_el.elements['variation']&.text&.strip
    
    slots = []
    tmpl_el.each_element('slot') { |s| slots << s }
    
    case selector
    when 'tmROOT'
      radicand_slot = slots.find { |s| s.elements['options']&.text&.strip != '1' } || slots[0]
      index_slot = slots.find { |s| s.elements['options']&.text&.strip == '1' }
      
      radicand = convert_slot(radicand_slot)
      
      if variation == 'tvROOT_SQ' || index_slot.nil? || slot_is_empty?(index_slot)
        "<msqrt><mrow>#{radicand}</mrow></msqrt>"
      else
        index = convert_slot(index_slot)
        "<mroot><mrow>#{radicand}</mrow><mrow>#{index}</mrow></mroot>"
      end
      
    when 'tmFRACT'
      num = slots[0] ? convert_slot(slots[0]) : ""
      den = slots[1] ? convert_slot(slots[1]) : ""
      "<mfrac><mrow>#{num}</mrow><mrow>#{den}</mrow></mfrac>"
      
    when 'tmSUP'
      base = slots[0] ? convert_slot(slots[0]) : ""
      exp = slots[1] ? convert_slot(slots[1]) : ""
      "<msup><mrow>#{base}</mrow><mrow>#{exp}</mrow></msup>"
      
    when 'tmSUB'
      base = slots[0] ? convert_slot(slots[0]) : ""
      sub = slots[1] ? convert_slot(slots[1]) : ""
      "<msub><mrow>#{base}</mrow><mrow>#{sub}</mrow></msub>"
      
    when 'tmSUBSUP'
      base = slots[0] ? convert_slot(slots[0]) : ""
      sub = slots[1] ? convert_slot(slots[1]) : ""
      sup = slots[2] ? convert_slot(slots[2]) : ""
      "<msubsup><mrow>#{base}</mrow><mrow>#{sub}</mrow><mrow>#{sup}</mrow></msubsup>"
      
    when 'tmPAREN'
      inner = slots[0] ? convert_slot(slots[0]) : ""
      "<mfenced open=\"(\" close=\")\"><mrow>#{inner}</mrow></mfenced>"
      
    when 'tmBRACK'
      inner = slots[0] ? convert_slot(slots[0]) : ""
      "<mfenced open=\"[\" close=\"]\"><mrow>#{inner}</mrow></mfenced>"
      
    when 'tmBRACE'
      inner = slots[0] ? convert_slot(slots[0]) : ""
      "<mfenced open=\"{\" close=\"}\"><mrow>#{inner}</mrow></mfenced>"
      
    when 'tmBAR', 'tmOBAR'
      inner = slots[0] ? convert_slot(slots[0]) : ""
      "<mover><mrow>#{inner}</mrow><mo>¯</mo></mover>"
      
    when 'tmVEC'
      inner = slots[0] ? convert_slot(slots[0]) : ""
      "<mover><mrow>#{inner}</mrow><mo>→</mo></mover>"
      
    when 'tmHAT'
      inner = slots[0] ? convert_slot(slots[0]) : ""
      "<mover><mrow>#{inner}</mrow><mo>^</mo></mover>"
      
    when 'tmLIM'
      if slots.length >= 2
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

# ============================================================
# Convert single file
# ============================================================
def convert_single(path)
  result = {
    path: path,
    mathml: nil,
    error: nil
  }

  unless File.exist?(path)
    result[:error] = "File not found"
    return result
  end

  begin
    # Get MTEF XML
    mtef_converter = Mathtype::Converter.new(path)
    mtef_xml = mtef_converter.to_xml
    
    has_sqrt = mtef_xml =~ /<selector>\s*tmROOT\s*<\/selector>/i
    
    mathml = ""
    
    if has_sqrt
      # Use custom converter for sqrt
      begin
        custom_converter = MtefToMathml.new(mtef_xml)
        mathml = custom_converter.convert || ""
        
        if mathml.empty? || !(mathml =~ /<msqrt|<mroot/i)
          if GEM_AVAILABLE
            converter = MathTypeToMathML::Converter.new(path)
            mathml = converter.convert || ""
          end
        end
      rescue => e
        if GEM_AVAILABLE
          converter = MathTypeToMathML::Converter.new(path)
          mathml = converter.convert || ""
        end
      end
    else
      # Use gem for non-sqrt
      if GEM_AVAILABLE
        converter = MathTypeToMathML::Converter.new(path)
        mathml = converter.convert || ""
      else
        custom_converter = MtefToMathml.new(mtef_xml)
        mathml = custom_converter.convert || ""
      end
    end
    
    result[:mathml] = mathml

  rescue => e
    result[:error] = "#{e.class}: #{e.message}"
  end

  result
end

# ============================================================
# Main: Process all files in batch
# ============================================================
if ARGV.empty?
  puts JSON.generate({ error: "No files provided" })
  exit 1
end

results = []

ARGV.each do |path|
  results << convert_single(path)
end

puts JSON.generate(results)
