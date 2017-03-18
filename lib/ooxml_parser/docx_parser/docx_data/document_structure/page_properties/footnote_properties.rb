module OoxmlParser
  # Class for parsing `footnotePr` tags
  class FootnoteProperties < OOXMLDocumentObject
    # @return [NumberingFormat] format of numbering
    attr_accessor :numbering_format
    # @return [ValuedChild] type of numbering restart
    attr_accessor :numbering_restart
    # @return [ValuedChild] value of numbering start
    attr_accessor :numbering_start
    # @return [ValuedChild] position of footnote
    attr_accessor :position

    # Parse FootnoteProperties
    # @param [Nokogiri::XML:Element] node with FootnoteProperties
    # @return [FootnoteProperties] value of FootnoteProperties
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'numFmt'
          @numbering_format = ValuedChild.new(:symbol, parent: self).parse(node_child)
        when 'numRestart'
          @numbering_restart = ValuedChild.new(:symbol, parent: self).parse(node_child)
        when 'numStart'
          @numbering_start = ValuedChild.new(:integer, parent: self).parse(node_child)
        when 'pos'
          @position = ValuedChild.new(:symbol, parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
