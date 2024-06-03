import tkinter as tk

class IdleGame:
    def __init__(self, root):
        self.root = root
        self.root.title("Idle Incremental Game")

        # Initialize points
        self.points = 0
        self.click_upgrade_cost = 5
        self.auto_upgrade_cost = 10
        self.points_per_click = 1
        self.auto_increment_value = 0
        self.auto_increment_rate = 1000  # in milliseconds

        # Create label to display points
        self.points_label = tk.Label(root, text=f"Points: {self.points}", font=('Helvetica', 16))
        self.points_label.pack(pady=10)

        # Create button to increment points
        self.increment_button = tk.Button(root, text="Click me!", command=self.increment_points, font=('Helvetica', 16))
        self.increment_button.pack(pady=10)

        # Upgrade shop
        self.shop_frame = tk.Frame(root)
        self.shop_frame.pack(pady=20)

        # buttons
        self.click_upgrade_button = tk.Button(self.shop_frame, text=f"Upgrade Click (+1) - Cost: {self.click_upgrade_cost}",
            command=self.upgrade_click, font=('Helvetica', 14))
        self.click_upgrade_button.pack(pady=5)

        self.auto_upgrade_button = tk.Button(self.shop_frame, text=f"Upgrade Auto (+0.5x) - Cost: {self.auto_upgrade_cost}",
            command=self.upgrade_auto, font=('Helvetica', 14))
        self.auto_upgrade_button.pack(pady=5)


        # Start auto-increment
        self.auto_increment()

        # Manual point gain button
    def increment_points(self):
        self.points += self.points_per_click
        self.update_points_label()

        # Auto-increment points every second
    def auto_increment(self):
        self.points += self.auto_increment_value
        self.update_points_label()
        self.root.after(int(self.auto_increment_rate), self.auto_increment)

    def update_points_label(self):
        self.points_label.config(text=f"Points: {self.points}")
        self.update_shop_buttons()

    def upgrade_click(self):
        if self.points >= self.click_upgrade_cost:
            self.points -= self.click_upgrade_cost
            self.points_per_click = int(self.points_per_click * 2)
            self.click_upgrade_cost = int(self.click_upgrade_cost * 1.5)  # Increase cost for next upgrade
            self.update_points_label()

            points_per_click = self.points_per_click
            print("Points Per Click:")
            print (points_per_click) 

    def upgrade_auto(self):  #Split this into seperate upgrades later
        if self.points >= self.auto_upgrade_cost:
            self.points -= self.auto_upgrade_cost
            self.auto_increment_value = int(self.auto_increment_value + 1)  # Increase auto-increment value
            self.auto_increment_rate = int(self.auto_increment_rate * 0.9)  # Decrease auto-increment interval
            self.auto_upgrade_cost = int(self.auto_upgrade_cost * 1.5)  # Increase cost for next upgrade
            self.update_points_label()

    def update_shop_buttons(self):
        self.click_upgrade_button.config(text=f"Upgrade Click (+1) - Cost: {self.click_upgrade_cost}")
        self.auto_upgrade_button.config(text=f"Upgrade Auto (+0.5x) - Cost: {self.auto_upgrade_cost}")

if __name__ == "__main__":
    root = tk.Tk()
    game = IdleGame(root)
    root.mainloop()
