# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <workbookPr> tag
  class WorkbookProperties < OOXMLDocumentObject
    # @return [True, False] Specifies if 1904 date system is used in workbook
    attr_reader :date1904

    def initialize(parent: nil)
      @date1904 = false
      super
    end

    # Parse WorkbookProperties data
    # @param [Nokogiri::XML:Element] node with WorkbookProperties data
    # @return [WorkbookProperties] value of WorkbookProperties
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'date1904'
          @date1904 = boolean_attribute_value(value)
        end
      end
      self
    end
  end
end
