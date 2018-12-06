require_relative 'pivot_formats/pivot_format'
module OoxmlParser
  # Parsing chart `c:pivotFmts`
  class PivotFormats < OOXMLDocumentObject
    # @return [Array<PivotFormat>] list of pivot formats
    attr_reader :pivot_formats_list

    def initialize(parent: nil)
      @parent = parent
      @pivot_formats_list = []
    end

    # @return [Array, PivotFormat] accessor
    def [](key)
      @pivot_formats_list[key]
    end

    # Parse PivotFormats object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [PivotFormats] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'pivotFmt'
          @pivot_formats_list << PivotFormat.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
