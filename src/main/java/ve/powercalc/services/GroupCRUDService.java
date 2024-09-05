package ve.powercalc.services;

import lombok.RequiredArgsConstructor;
import ve.powercalc.entity.Group;
import ve.powercalc.repositories.GroupRepository;
import org.springframework.stereotype.Service;
import java.util.Collection;


@RequiredArgsConstructor
@Service
public class GroupCRUDService implements CRUDService <Group> {

    private final GroupRepository repository;

    @Override
    public Group getById(Integer id) {
        return repository.findById(id).orElseThrow();
    }

    @Override
    public Collection<Group> getAll() {
        return repository.findAll();
    }

    @Override
    public void create(Group item) {
        repository.save(item);
    }

    @Override
    public void update(Group item) {
        repository.save(item);
    }

    @Override
    public void delete(Integer id) {
        repository.deleteById(id);
    }

}
