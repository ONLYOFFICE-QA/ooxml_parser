# frozen_string_literal: true

module OoxmlParser
  # Class for `dxf` data
  class DifferentialFormattingRecord < OOXMLDocumentObject
    # @return [Font] Font
    attr_reader :font
    # @return [NumberFormat] Number format
    attr_reader :number_format
    # @return [Fill] Fill
    attr_reader :fill
    # @return [Borders] Borders
    attr_reader :borders

    # Parse DifferentialFormattingRecord data
    # @param [Nokogiri::XML:Element] node with DifferentialFormattingRecord data
    # @return [DifferentialFormattingRecord] value of DifferentialFormattingRecord data
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'font'
          @font = Font.new(parent: self).parse(node_child)
        when 'numFmt'
          @number_format = NumberFormat.new(parent: self).parse(node_child)
        when 'fill'
          @fill = Fill.new(parent: self).parse(node_child)
        when 'border'
          @borders = XlsxBorder.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
