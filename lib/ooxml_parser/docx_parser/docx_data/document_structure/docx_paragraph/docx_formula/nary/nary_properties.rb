require_relative 'nary_properties/nary_grow'
require_relative 'nary_properties/nary_limit_location'
module OoxmlParser
  # Class for parsing `m:naryPr` object
  class NaryProperties
    # @return [NaryGrow] grow value
    attr_accessor :grow
    # @return [NaryLimitLocation] nary limit location
    attr_accessor :limit_location

    # Parse NaryProperties
    # @param [Nokogiri::XML:Node] node with NaryProperties
    # @return [NaryProperties] result of parsing
    def self.parse(node)
      nary_properties = NaryProperties.new
      node.xpath('*').each do |nary_pr_child|
        case nary_pr_child.name
        when 'grow'
          nary_properties.grow = NaryGrow.parse(nary_pr_child)
        when 'limLoc'
          nary_properties.limit_location = NaryLimitLocation.new(parent: nary_properties).parse(nary_pr_child)
        end
      end
      nary_properties
    end
  end
end
