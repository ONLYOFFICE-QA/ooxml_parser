require_relative 'styles/document_defaults'
module OoxmlParser
  # Class for parsing `styles.xml` file
  class Styles < OOXMLDocumentObject
    # @return [DocumentDefaults] defaults of document
    attr_accessor :document_defaults
    # @return [Array<DocumentStyle>] array of document styles
    attr_reader :styles

    def initialize(parent: nil)
      @styles = []
      @parent = parent
    end

    def parse
      doc = parse_xml(OOXMLDocumentObject.path_to_folder + 'word/styles.xml')
      doc.xpath('w:styles/*').each do |node_child|
        case node_child.name
        when 'docDefaults'
          @document_defaults = DocumentDefaults.new(parent: self).parse(node_child)
        when 'style'
          @styles << DocumentStyle.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
