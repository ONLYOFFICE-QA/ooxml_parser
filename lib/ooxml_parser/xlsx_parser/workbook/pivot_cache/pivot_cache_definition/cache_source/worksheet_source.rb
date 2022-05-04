# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <worksheetSource> file
  class WorksheetSource < OOXMLDocumentObject
    # @return [String] ref of cells
    attr_reader :ref
    # @return [String] sheet name
    attr_reader :sheet

    # Parse `<worksheetSource>` tag
    # # @param [Nokogiri::XML:Element] node with WorksheetSource data
    # @return [WorksheetSource]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'ref'
          @ref = value.value.to_s
        when 'sheet'
          @sheet = value.value.to_s
        end
      end
      self
    end
  end
end
