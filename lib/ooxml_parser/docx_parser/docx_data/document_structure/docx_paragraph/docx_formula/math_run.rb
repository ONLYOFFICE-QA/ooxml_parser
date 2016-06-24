require_relative 'math_run/math_run_properties'
module OoxmlParser
  # Class for parsing `m:r` object
  class MathRun
    # @return [MathRunProperties] properties of run
    attr_accessor :properties
    # @return [String] text of formula
    attr_accessor :text

    # Parse MathRun
    # @param [Nokogiri::XML:Node] node with MathRun
    # @return [MathRun] result of parsing
    def self.parse(node)
      math_run = MathRun.new
      node.xpath('*').each do |math_run_child|
        case math_run_child.name
        when 'rPr'
          math_run.properties = MathRunProperties.parse(math_run_child)
        when 't'
          math_run.text = math_run_child.text
        end
      end
      math_run
    end
  end
end
