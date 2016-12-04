require_relative 'nary_properties/nary_grow'
require_relative 'nary_properties/nary_limit_location'
module OoxmlParser
  # Class for parsing `m:naryPr` object
  class NaryProperties < OOXMLDocumentObject
    # @return [NaryGrow] grow value
    attr_accessor :grow
    # @return [NaryLimitLocation] nary limit location
    attr_accessor :limit_location

    # Parse NaryProperties
    # @param [Nokogiri::XML:Node] node with NaryProperties
    # @return [NaryProperties] result of parsing
    def parse(node)
      node.xpath('*').each do |nary_pr_child|
        case nary_pr_child.name
        when 'grow'
          @grow = NaryGrow.new(parent: self).parse(nary_pr_child)
        when 'limLoc'
          @limit_location = NaryLimitLocation.new(parent: self).parse(nary_pr_child)
        end
      end
      self
    end
  end
end
