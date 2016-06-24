module OoxmlParser
  # Class for parsing `m:rPr` object
  class MathRunProperties
    # @return [True, False] is run with break
    attr_accessor :break

    def initialize
      @break = false
    end

    # Parse MathRunProperties
    # @param [Nokogiri::XML:Node] node with MathRunProperties
    # @return [MathRunProperties] result of parsing
    def self.parse(node)
      math_run_properties = MathRunProperties.new
      node.xpath('*').each do |math_run_child|
        case math_run_child.name
        when 'brk'
          math_run_properties.break = true
        end
      end
      math_run_properties
    end
  end
end
