# mt2mml_v2.rb - Custom converter with proper XML parsing for tmROOT
require 'json'

# Load required gems with error handling
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

path = ARGV[0]
unless path && File.exist?(path)
  puts JSON.generate({ error: "File not found: #{path}" })
  exit 1
end

result = {
  mtef_xml: nil,
  mathml: nil,
  has_sqrt_mtef: false,
  has_sqrt_mathml: false,
  sqrt_selectors: [],
  sqrt_fixed: false,
  fix_method: nil,
  error: nil
}

# ============================================================
# Custom MTEF XML to MathML converter using REXML
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
    
    # Skip problematic characters
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
# Main conversion logic
# ============================================================

begin
  # Step 1: Get MTEF XML
  mtef_converter = Mathtype::Converter.new(path)
  result[:mtef_xml] = mtef_converter.to_xml
  mtef_xml = result[:mtef_xml] || ""
  
  # Check for sqrt in MTEF XML
  result[:has_sqrt_mtef] = true if mtef_xml =~ /<selector>\s*tmROOT\s*<\/selector>/i
  
  # Capture all selector values
  mtef_xml.scan(/<selector>([^<]+)<\/selector>/i).each do |match|
    selector = match[0].strip
    result[:sqrt_selectors] << selector unless result[:sqrt_selectors].include?(selector)
  end
  
  # Step 2: Choose converter
  mathml = ""
  
  if result[:has_sqrt_mtef]
    # Formula HAS sqrt - use custom converter (gem is broken for tmROOT)
    begin
      custom_converter = MtefToMathml.new(mtef_xml)
      mathml = custom_converter.convert || ""
      
      if mathml =~ /<msqrt|<mroot/i
        result[:sqrt_fixed] = true
        result[:fix_method] = "custom_mtef_converter"
      elsif GEM_AVAILABLE
        converter = MathTypeToMathML::Converter.new(path)
        mathml = converter.convert || ""
        result[:fix_method] = "gem_fallback_for_sqrt"
      end
    rescue => e
      STDERR.puts "Custom converter error: #{e.message}"
      if GEM_AVAILABLE
        converter = MathTypeToMathML::Converter.new(path)
        mathml = converter.convert || ""
        result[:fix_method] = "gem_fallback_error"
      end
    end
  else
    # Formula has NO sqrt - use gem (works fine)
    if GEM_AVAILABLE
      begin
        converter = MathTypeToMathML::Converter.new(path)
        mathml = converter.convert || ""
        result[:fix_method] = "gem"
      rescue => e
        STDERR.puts "Gem error: #{e.message}"
        # Fallback to custom
        begin
          custom_converter = MtefToMathml.new(mtef_xml)
          mathml = custom_converter.convert || ""
          result[:fix_method] = "custom_fallback"
        rescue => e2
          STDERR.puts "Both converters failed: #{e2.message}"
        end
      end
    else
      # No gem, use custom
      begin
        custom_converter = MtefToMathml.new(mtef_xml)
        mathml = custom_converter.convert || ""
        result[:fix_method] = "custom_only"
      rescue => e
        STDERR.puts "Custom converter error: #{e.message}"
      end
    end
  end
  
  result[:mathml] = mathml
  result[:has_sqrt_mathml] = (mathml =~ /<msqrt|<mroot/i) ? true : false

rescue => e
  result[:error] = "#{e.class}: #{e.message}"
  STDERR.puts "Main error: #{e.class}: #{e.message}"
  STDERR.puts e.backtrace.first(5).join("\n")
end

puts JSON.generate(result)
