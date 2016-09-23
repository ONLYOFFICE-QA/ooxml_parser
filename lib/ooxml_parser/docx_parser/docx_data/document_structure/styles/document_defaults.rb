require_relative 'document_defaults/paragraph_properties_default'
require_relative 'document_defaults/run_properties_default'
module OoxmlParser
  # Class for parsing `w:docDefaults` tags
  class DocumentDefaults < OOXMLDocumentObject
    # @return [RunPropertiesDefault] default properties of run
    attr_accessor :run_properties_default
    # @return [RunPropertiesDefault] default properties of run
    attr_accessor :paragraph_properties_default

    # Parse Bookmark object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Bookmark] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'rPrDefault'
          @run_properties_default = RunPropertiesDefault.new(parent: self).parse(node_child)
        when 'pPrDefault'
          @paragraph_properties_default = ParagraphPropertiesDefault.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
