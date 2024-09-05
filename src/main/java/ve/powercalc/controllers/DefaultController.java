package ve.powercalc.controllers;


import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import ve.powercalc.PowerCount;
import ve.powercalc.entity.Group;
import ve.powercalc.services.GroupCRUDService;
import java.io.File;
import java.io.IOException;
import java.net.URLDecoder;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.SQLException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Collection;

@Controller
@RequestMapping("/")
public class DefaultController {
    private final GroupCRUDService groupService;

    public DefaultController(GroupCRUDService groupService) {
        this.groupService = groupService;
    }

    /**
     * Метод формирует страницу из HTML-файла index.html,
     * который находится в папке resources/templates.
     * Это делает библиотека Thymeleaf.
     */

    @GetMapping ("/")
    public String index(Model model) {
        Collection<Group> groups = groupService.getAll();
        model.addAttribute("groups", groups);
        String dateFrom = LocalDateTime.now().minusMonths(1).format(DateTimeFormatter.ofPattern("yyyy-MM-01"));
        String dateTo = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-01"));
        model.addAttribute("dateFrom", dateFrom);
        model.addAttribute("dateTo", dateTo);
        return "index";
    }

    @PostMapping("/")
    public ResponseEntity<Resource> getPower(@RequestParam String group, @ModelAttribute("dateFrom") String dateFrom, @ModelAttribute("dateTo") String dateTo) throws SQLException, IOException {
        String filename = "/home/NW.MRSKSEVZAP.RU/vol00345/Документы/Расчет за мощность " +
                LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd-HH-mm-ss")) + ".xls" ;

        Integer id = Integer.parseInt(group);
        PowerCount.getPower(id, dateFrom, dateTo, filename);

        return downloadFile(filename);
    }

    public ResponseEntity<Resource> downloadFile(String filename) throws IOException {
        File file = new File(filename);
        Path filePath = Paths.get(file.getAbsolutePath());
        ByteArrayResource fileResource = new ByteArrayResource(Files.readAllBytes(filePath));

        String name = URLEncoder.encode(file.getName(), StandardCharsets.UTF_8);
        name = URLDecoder.decode(name, "ISO8859_1");

        String disposition = "inline; filename=" + name;
        HttpHeaders headers = new HttpHeaders();
        headers.add(HttpHeaders.CONTENT_DISPOSITION, disposition);

        return ResponseEntity.ok()
                .headers(headers)
                .contentLength(file.length())
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(fileResource);
    }
}

