# frozen_string_literal: true

module OoxmlParser
  # Parsing Chart 3D view tag 'view3D'
  class View3D < OOXMLDocumentObject
    # @return [ValuedChild] This element specifies the amount
    # a 3-D chart shall be rotated in the X direction
    attr_accessor :rotation_x
    # @return [ValuedChild] This element specifies the amount
    # a 3-D chart shall be rotated in the Y direction
    attr_accessor :rotation_y
    # @return [True, False] This element specifies that the
    # chart axes are at right angles, rather than drawn
    # in perspective. Applies only to 3-D charts.
    attr_accessor :right_angle_axis

    # Parse View3D object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [View3D] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'rotX'
          @rotation_x = ValuedChild.new(:integer, parent: self).parse(node_child)
        when 'rotY'
          @rotation_y = ValuedChild.new(:integer, parent: self).parse(node_child)
        when 'rAngAx'
          @right_angle_axis = option_enabled?(node_child)
        end
      end
      self
    end
  end
end
