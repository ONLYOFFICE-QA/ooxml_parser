# frozen_string_literal: true

module OoxmlParser
  # Class for storing data of chartsheets files
  class Chartsheet < OOXMLDocumentObject
    # @return [Array, SheetView] array of views
    attr_accessor :sheet_views

    def initialize(parent: nil)
      @sheet_views = []
      super
    end

    # Parse Chartsheet object
    # @param file [String] file to parse
    # @return [Chartsheet] result of parsing
    def parse(file)
      OOXMLDocumentObject.add_to_xmls_stack(OOXMLDocumentObject.root_subfolder + file)
      doc = parse_xml(OOXMLDocumentObject.current_xml)
      node = doc.xpath('//xmlns:chartsheet').first
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'sheetViews'
          node_child.xpath('*').each do |view_child|
            @sheet_views << SheetView.new(parent: self).parse(view_child)
          end
        end
      end
      OOXMLDocumentObject.xmls_stack.pop
      self
    end
  end
end
