require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/extensions'

module RubyXL
  # http://www.datypic.com/sc/ooxml/e-ssml_c-1.html
  class CalculationChainCell < OOXMLObject
    define_attribute(:r, :ref,  :accessor => :ref)
    define_attribute(:i, :int,  :accessor => :sheet_id,    :default => 0)
    define_attribute(:s, :bool, :accessor => :child_chain, :default => false)
    define_attribute(:l, :bool, :accessor => :new_dep_lvl, :default => false)
    define_attribute(:t, :bool, :accessor => :new_thread,  :default => false)
    define_attribute(:a, :bool, :accessor => :array,       :default => false)
    define_element_name 'c'
  end

  # http://www.datypic.com/sc/ooxml/e-ssml_calcChain.html
  class CalculationChain < OOXMLTopLevelObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain'

    define_child_node(RubyXL::CalculationChainCell, :collection => true, :accessor => :cells)
    define_child_node(RubyXL::ExtensionStorageArea)

    define_element_name 'calcChain'
    set_namespaces('http://schemas.openxmlformats.org/spreadsheetml/2006/main' => nil)

    def xlsx_path
      ROOT.join('xl', 'calcChain.xml')
    end

    def update_after(action, position, direction = nil)
      row      = position[0]
      col      = position[1]
      sheet_id = position[2]

      case action
      when :insert_row
        cells.each { |calc_chain_cell|
          next if calc_chain_cell.sheet_id != sheet_id
          ref = calc_chain_cell.ref
          next if ref.first_row < row
          calc_chain_cell.ref = RubyXL::Reference.new(ref.first_row + 1, ref.first_col)
        }
      when :delete_row
        calc_chain_cells_to_remove = []

        cells.each { |calc_chain_cell|
          next if calc_chain_cell.sheet_id != sheet_id
          ref = calc_chain_cell.ref
          next if ref.first_row < row

          if ref.first_row == row
            calc_chain_cells_to_remove << calc_chain_cell
          else
            calc_chain_cell.ref = RubyXL::Reference.new(ref.first_row - 1, ref.first_col)
          end
        }

        cells.reject! { |calc_chain_cell| calc_chain_cells_to_remove.include?(calc_chain_cell) }
      when :insert_column
        cells.each { |calc_chain_cell|
          next if calc_chain_cell.sheet_id != sheet_id
          ref = calc_chain_cell.ref
          next if ref.first_col < col
          calc_chain_cell.ref = RubyXL::Reference.new(ref.first_row, ref.first_col + 1)
        }
      when :delete_column
        calc_chain_cells_to_remove = []

        cells.each { |calc_chain_cell|
          next if calc_chain_cell.sheet_id != sheet_id
          ref = calc_chain_cell.ref
          next if ref.first_col < col

          if ref.first_col == col
            calc_chain_cells_to_remove << calc_chain_cell
          else
            calc_chain_cell.ref = RubyXL::Reference.new(ref.first_row, ref.first_col - 1)
          end
        }

        cells.reject! { |calc_chain_cell| calc_chain_cells_to_remove.include?(calc_chain_cell) }
      when :insert_cell
        if direction == :right
          cells.each { |calc_chain_cell|
            next if calc_chain_cell.sheet_id != sheet_id
            ref = calc_chain_cell.ref
            next if ref.first_row != row || ref.first_col < col
            calc_chain_cell.ref = RubyXL::Reference.new(ref.first_row, ref.first_col + 1)
          }
        elsif direction == :down
          cells.each { |calc_chain_cell|
            next if calc_chain_cell.sheet_id != sheet_id
            ref = calc_chain_cell.ref
            next if ref.first_row < row || ref.first_col != col
            calc_chain_cell.ref = RubyXL::Reference.new(ref.first_row + 1, ref.first_col)
          }
        end
      when :delete_cell
        calc_chain_cells_to_remove = []

        if direction == :left
          cells.each { |calc_chain_cell|
            next if calc_chain_cell.sheet_id != sheet_id
            ref = calc_chain_cell.ref
            next if ref.first_row != row || ref.first_col < col

            if ref.first_col == col
              calc_chain_cells_to_remove << calc_chain_cell
            else
              calc_chain_cell.ref = RubyXL::Reference.new(ref.first_row, ref.first_col - 1)
            end
          }
        elsif direction == :up
          cells.each { |calc_chain_cell|
            next if calc_chain_cell.sheet_id != sheet_id
            ref = calc_chain_cell.ref
            next if ref.first_row < row || ref.first_col != col

            if ref.first_row == row
              calc_chain_cells_to_remove << calc_chain_cell
            else
              calc_chain_cell.ref = RubyXL::Reference.new(ref.first_row - 1, ref.first_col)
            end
          }
        end

        cells.reject! { |calc_chain_cell| calc_chain_cells_to_remove.include?(calc_chain_cell) }
      end
    end
  end
end
