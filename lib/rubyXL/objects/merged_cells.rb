module RubyXL
  class MergedCells

    def update_after(action, position, direction = nil)
      row = position[0]
      col = position[1]

      case action
      when :insert_row
        each { |merged_cell|
          next if merged_cell.ref.last_row < row
          ref = merged_cell.ref
          first_row = ref.first_row
          first_row += 1 if first_row >= row
          merged_cell.ref = RubyXL::Reference.new(first_row, ref.last_row + 1, ref.first_col, ref.last_col)
        }
      when :delete_row
        each { |merged_cell|
          next if merged_cell.ref.last_row <= row
          ref = merged_cell.ref
          first_row = ref.first_row
          first_row -= 1 if first_row > row
          merged_cell.ref = RubyXL::Reference.new(first_row, ref.last_row - 1, ref.first_col, ref.last_col)
        }
      when :insert_column
        each { |merged_cell|
          next if merged_cell.ref.last_col < col
          ref = merged_cell.ref
          first_col = ref.first_col
          first_col += 1 if first_col >= col
          merged_cell.ref = RubyXL::Reference.new(ref.first_row, ref.last_row, first_col, ref.last_col + 1)
        }
      when :delete_column
        each { |merged_cell|
          next if merged_cell.ref.last_col <= col
          ref = merged_cell.ref
          first_col = ref.first_col
          first_col -= 1 if first_col > col
          merged_cell.ref = RubyXL::Reference.new(ref.first_row, ref.last_row, first_col, ref.last_col - 1)
        }
      when :insert_cell
        cells_to_unmerge = []

        if direction == :right
          each { |merged_cell|
            ref = merged_cell.ref
            next if ref.first_row > row || ref.last_row < row || ref.last_col < col
            if ref.first_row == ref.last_row
              first_col = ref.first_col
              first_col += 1 if first_col >= col
              merged_cell.ref = RubyXL::Reference.new(ref.first_row, ref.last_row, first_col, ref.last_col + 1)
            else
              cells_to_unmerge << merged_cell
            end
          }
        elsif direction == :down
          each { |merged_cell|
            ref = merged_cell.ref
            next if ref.first_col > col || ref.last_col < col || ref.last_row < row
            if ref.first_col == ref.last_col
              first_row = ref.first_row
              first_row += 1 if first_row >= row
              merged_cell.ref = RubyXL::Reference.new(first_row, ref.last_row + 1, ref.first_col, ref.last_col)
            else
              cells_to_unmerge << merged_cell
            end
          }
        end

        merged_cells.reject!{ |merged_cell| cells_to_unmerge.include?(merged_cell) }
      when :delete_cell
        cells_to_unmerge = []

        if direction == :left
          each { |merged_cell|
            ref = merged_cell.ref
            next if ref.first_row > row || ref.last_row < row || ref.last_col < col
            if ref.first_row == ref.last_row
              first_col = ref.first_col
              first_col -= 1 if first_col > col
              merged_cell.ref = RubyXL::Reference.new(ref.first_row, ref.last_row, first_col, ref.last_col - 1)
            else
              cells_to_unmerge << merged_cell
            end
          }
        elsif direction == :up
          each { |merged_cell|
            ref = merged_cell.ref
            next if ref.first_col > col || ref.last_col < col || ref.last_row < row
            if ref.first_col == ref.last_col
              first_row = ref.first_row
              first_row -= 1 if first_row > row
              merged_cell.ref = RubyXL::Reference.new(first_row, ref.last_row - 1, ref.first_col, ref.last_col)
            else
              cells_to_unmerge << merged_cell
            end
          }
        end

        merged_cells.reject!{ |merged_cell| cells_to_unmerge.include?(merged_cell) }
      end
    end
  end
end
