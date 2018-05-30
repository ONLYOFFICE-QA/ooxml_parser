require_relative 'cell_style/alignment'
module OoxmlParser
  # Class for parsing cell style
  class CellStyle < OOXMLDocumentObject
    # Parse CellStyle object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [CellStyle] result of parsing
    def parse(node)
      xf = root_object.style_sheet.cell_xfs.xf_array[node.to_i]
      xf.calculate_values
    end
  end
end
